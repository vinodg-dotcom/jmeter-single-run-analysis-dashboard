"""
Single-Run JMeter Load Test Analysis Dashboard
A Streamlit app for analysing one JMeter CSV/JTL results file (no baseline comparison).
"""

from __future__ import annotations

from io import BytesIO
from datetime import datetime
import math

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as ExcelImage
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

st.set_page_config(
    page_title="Load Test Results & Analysis",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📈 Load Test Results & Analysis")
st.caption("Single-run Apache JMeter analysis · response times · throughput · errors · concurrency")

REQUIRED_ALIASES = {
    "timestamp": ["timeStamp", "timestamp", "time_stamp", "start_time"],
    "elapsed_ms": ["elapsed", "elapsed_ms", "response_time", "response time"],
    "transaction": ["label", "transaction", "sampler", "name"],
    "response_code": ["responseCode", "response_code", "status_code", "status"],
    "success": ["success", "successful", "is_success"],
    "threads": ["allThreads", "all_threads", "threads", "active_threads"],
    "latency_ms": ["Latency", "latency", "latency_ms"],
    "connect_ms": ["Connect", "connect", "connect_ms"],
    "received_bytes": ["bytes", "received_bytes", "bodySize"],
    "sent_bytes": ["sentBytes", "sent_bytes"],
    "url": ["URL", "url", "uri"],
}


def find_column(df: pd.DataFrame, aliases: list[str]) -> str | None:
    lookup = {str(c).strip().lower(): c for c in df.columns}
    for alias in aliases:
        if alias.strip().lower() in lookup:
            return lookup[alias.strip().lower()]
    return None


def to_bool_success(series: pd.Series) -> pd.Series:
    values = series.astype(str).str.strip().str.lower()
    return values.isin(["true", "1", "yes", "y", "success", "passed"])


def parse_results(uploaded) -> tuple[pd.DataFrame, list[str]]:
    raw = pd.read_csv(uploaded, low_memory=False)
    if raw.empty:
        raise ValueError("The uploaded file contains no result rows.")

    mapped: dict[str, str] = {}
    for target, aliases in REQUIRED_ALIASES.items():
        source = find_column(raw, aliases)
        if source:
            mapped[target] = source

    missing = [x for x in ["timestamp", "elapsed_ms", "transaction"] if x not in mapped]
    if missing:
        raise ValueError(
            "Missing required JMeter columns. The file must contain timestamp (timeStamp), "
            "elapsed (elapsed), and transaction name (label)."
        )

    df = pd.DataFrame(index=raw.index)
    df["timestamp_raw"] = raw[mapped["timestamp"]]
    df["elapsed_ms"] = pd.to_numeric(raw[mapped["elapsed_ms"]], errors="coerce")
    df["transaction"] = raw[mapped["transaction"]].fillna("Unlabelled").astype(str)
    df["response_code"] = (
        raw[mapped["response_code"]].fillna("Unknown").astype(str)
        if "response_code" in mapped else "Unknown"
    )
    if "success" in mapped:
        df["success"] = to_bool_success(raw[mapped["success"]])
    elif "response_code" in mapped:
        df["success"] = df["response_code"].str.match(r"^2\d\d$")
    else:
        df["success"] = True

    df["threads"] = (
        pd.to_numeric(raw[mapped["threads"]], errors="coerce")
        if "threads" in mapped else np.nan
    )
    df["latency_ms"] = (
        pd.to_numeric(raw[mapped["latency_ms"]], errors="coerce")
        if "latency_ms" in mapped else np.nan
    )
    df["connect_ms"] = (
        pd.to_numeric(raw[mapped["connect_ms"]], errors="coerce")
        if "connect_ms" in mapped else np.nan
    )
    df["received_bytes"] = (
        pd.to_numeric(raw[mapped["received_bytes"]], errors="coerce")
        if "received_bytes" in mapped else np.nan
    )
    df["sent_bytes"] = (
        pd.to_numeric(raw[mapped["sent_bytes"]], errors="coerce")
        if "sent_bytes" in mapped else np.nan
    )
    df["url"] = raw[mapped["url"]].fillna("").astype(str) if "url" in mapped else ""

    # JMeter's timeStamp is normally Unix epoch milliseconds. Also accept parseable date strings.
    numeric_ts = pd.to_numeric(df["timestamp_raw"], errors="coerce")
    if numeric_ts.notna().mean() > 0.8:
        unit = "ms" if numeric_ts.dropna().median() > 10**11 else "s"
        df["start_time"] = pd.to_datetime(numeric_ts, unit=unit, errors="coerce", utc=True)
    else:
        df["start_time"] = pd.to_datetime(df["timestamp_raw"], errors="coerce", utc=True)

    df["elapsed_ms"] = df["elapsed_ms"].clip(lower=0)
    df = df.dropna(subset=["start_time", "elapsed_ms", "transaction"]).copy()
    if df.empty:
        raise ValueError("No valid rows remain after parsing timestamps and elapsed times.")

    df["end_time"] = df["start_time"] + pd.to_timedelta(df["elapsed_ms"], unit="ms")
    df["elapsed_seconds"] = df["elapsed_ms"] / 1000.0
    df["failed"] = ~df["success"]
    df["result"] = np.where(df["success"], "Success", "Failure")
    df["relative_seconds"] = (df["start_time"] - df["start_time"].min()).dt.total_seconds()
    warnings = []
    if "success" not in mapped and "response_code" not in mapped:
        warnings.append("Neither success nor responseCode was found; rows are assumed successful. Error rate may be inaccurate.")
    if "threads" not in mapped:
        warnings.append("No allThreads/concurrency column found; concurrency charts will be unavailable.")
    if df["response_code"].eq("Unknown").all():
        warnings.append("No HTTP response-code column found; response-code breakdown is unavailable.")
    return df.sort_values("start_time").reset_index(drop=True), warnings


def build_transaction_summary(df: pd.DataFrame) -> pd.DataFrame:
    grouped = df.groupby("transaction", dropna=False)
    summary = grouped.agg(
        samples=("elapsed_ms", "size"),
        average_ms=("elapsed_ms", "mean"),
        median_ms=("elapsed_ms", "median"),
        min_ms=("elapsed_ms", "min"),
        max_ms=("elapsed_ms", "max"),
        std_dev_ms=("elapsed_ms", "std"),
        failures=("failed", "sum"),
        success_rate_pct=("success", lambda s: 100 * s.mean()),
    )
    pcts = grouped["elapsed_ms"].quantile([0.90, 0.95, 0.99]).unstack()
    pcts.columns = ["p90_ms", "p95_ms", "p99_ms"]
    summary = summary.join(pcts)
    duration = max((df["end_time"].max() - df["start_time"].min()).total_seconds(), 0.001)
    summary["transaction_rate_req_s"] = summary["samples"] / duration
    summary["error_rate_pct"] = 100 * summary["failures"] / summary["samples"]
    return summary.reset_index().sort_values("average_ms", ascending=False)


def make_excel(df: pd.DataFrame, summary: pd.DataFrame, status_summary: pd.DataFrame, run_summary: pd.DataFrame) -> bytes:
    """Build an Excel workbook with all dashboard graphs embedded on a Charts sheet."""
    output = BytesIO()

    def excel_safe(frame: pd.DataFrame) -> pd.DataFrame:
        safe = frame.copy()
        for column in safe.columns:
            if isinstance(safe[column].dtype, pd.DatetimeTZDtype):
                safe[column] = safe[column].dt.tz_localize(None)
            elif safe[column].dtype == object:
                safe[column] = safe[column].map(
                    lambda value: value.tz_localize(None)
                    if isinstance(value, pd.Timestamp) and value.tzinfo is not None
                    else value
                )
        return safe

    charts: list[tuple[str, bytes]] = []

    def capture_chart(title: str, draw_fn):
        fig, ax = plt.subplots(figsize=(9, 4.8))
        try:
            draw_fn(fig, ax)
            ax.set_title(title, fontsize=13, pad=12)
            fig.tight_layout()
            image_buffer = BytesIO()
            fig.savefig(image_buffer, format="png", dpi=130, bbox_inches="tight")
            charts.append((title, image_buffer.getvalue()))
        finally:
            plt.close(fig)

    tx = summary.copy()
    tx["transaction"] = tx["transaction"].astype(str)
    tx_ranked = tx.sort_values("average_ms", ascending=False).head(10).sort_values("average_ms")

    capture_chart("Average vs P95 response time by transaction", lambda fig, ax: (
        ax.barh(tx_ranked["transaction"], tx_ranked["average_ms"], label="Average"),
        ax.barh(tx_ranked["transaction"], tx_ranked["p95_ms"], alpha=0.65, label="P95"),
        ax.set_xlabel("Response time (ms)"), ax.legend()
    ))
    capture_chart("P95 response time by transaction", lambda fig, ax: (
        ax.barh(tx_ranked.sort_values("p95_ms")["transaction"], tx_ranked.sort_values("p95_ms")["p95_ms"]),
        ax.set_xlabel("P95 (ms)")
    ))

    def draw_box(fig, ax):
        labels, values = [], []
        for name, part in df.groupby("transaction"):
            if len(part):
                labels.append(str(name))
                values.append(part["elapsed_ms"].dropna().to_numpy())
        if values:
            ax.boxplot(values, labels=labels, showfliers=True)
            ax.set_ylabel("Response time (ms)")
            ax.tick_params(axis="x", labelrotation=35, labelsize=8)
        else:
            ax.text(0.5, 0.5, "No response-time data", ha="center")
    capture_chart("Response-time distribution by transaction (box plot)", draw_box)

    capture_chart("Individual response times over time (scatter)", lambda fig, ax: (
        [ax.scatter(part["relative_seconds"], part["elapsed_ms"], s=12, alpha=0.65, label=str(name))
         for name, part in df.groupby("transaction")],
        ax.set_xlabel("Seconds since first sample"), ax.set_ylabel("Response time (ms)"),
        ax.legend(fontsize=7, loc="best")
    ))

    duration = max((df["end_time"].max() - df["start_time"].min()).total_seconds(), 0.001)
    rel = df["relative_seconds"].fillna(0)
    buckets = (np.floor(rel / 1).astype(int))
    throughput = df.assign(time_bucket=buckets).groupby("time_bucket").size()
    capture_chart("Requests per second over time (1-second buckets)", lambda fig, ax: (
        ax.plot(throughput.index, throughput.values, marker="o", markersize=3),
        ax.set_xlabel("Seconds since first sample"), ax.set_ylabel("Requests / second")
    ))

    capture_chart("Cumulative completed samples", lambda fig, ax: (
        ax.plot(rel.sort_values().to_numpy(), np.arange(1, len(df) + 1)),
        ax.set_xlabel("Seconds since first sample"), ax.set_ylabel("Cumulative samples")
    ))

    if df["threads"].notna().any():
        threads = df[["relative_seconds", "threads"]].dropna().sort_values("relative_seconds")
        capture_chart("Active threads recorded with samples", lambda fig, ax: (
            ax.plot(threads["relative_seconds"], threads["threads"], marker=".", markersize=3),
            ax.set_xlabel("Seconds since first sample"), ax.set_ylabel("Active threads")
        ))

    capture_chart("Sample outcome distribution", lambda fig, ax: (
        ax.pie([int(df["success"].sum()), int((~df["success"]).sum())],
               labels=["Successful", "Failed"], autopct="%1.1f%%", startangle=90)
    ))

    status = status_summary.copy()
    capture_chart("Samples by HTTP response code", lambda fig, ax: (
        ax.bar(status["response_code"].astype(str), status["samples"]),
        ax.set_xlabel("Response code"), ax.set_ylabel("Samples")
    ))

    failures = df.groupby("transaction").agg(
        samples=("transaction", "size"), failures=("failed", "sum")
    ).reset_index()
    failures["error_rate_pct"] = 100 * failures["failures"] / failures["samples"]
    failures = failures.sort_values("error_rate_pct", ascending=False).head(10)
    capture_chart("Error rate by transaction", lambda fig, ax: (
        ax.bar(failures["transaction"].astype(str), failures["error_rate_pct"]),
        ax.set_ylabel("Error rate (%)"), ax.tick_params(axis="x", labelrotation=35, labelsize=8)
    ))

    if df["latency_ms"].notna().any():
        latency_df = df.dropna(subset=["latency_ms"])
        capture_chart("Server latency by transaction", lambda fig, ax: (
            ax.boxplot([part["latency_ms"].to_numpy() for _, part in latency_df.groupby("transaction")],
                       labels=[str(name) for name, _ in latency_df.groupby("transaction")],
                       showfliers=True),
            ax.set_ylabel("Latency (ms)"), ax.tick_params(axis="x", labelrotation=35, labelsize=8)
        ))
    if df["connect_ms"].notna().any():
        connect_values = df["connect_ms"].dropna()
        capture_chart("Connection-time distribution", lambda fig, ax: (
            ax.hist(connect_values, bins=30),
            ax.set_xlabel("Connect time (ms)"), ax.set_ylabel("Samples")
        ))

    with pd.ExcelWriter(output, engine="openpyxl", datetime_format="yyyy-mm-dd hh:mm:ss.000") as writer:
        excel_safe(run_summary).to_excel(writer, index=False, sheet_name="Run Summary")
        excel_safe(summary).to_excel(writer, index=False, sheet_name="Transaction Summary")
        excel_safe(status_summary).to_excel(writer, index=False, sheet_name="Status Codes")
        excel_safe(df).to_excel(writer, index=False, sheet_name="All Samples")
        chart_ws = writer.book.create_sheet("Charts")
        chart_ws["A1"] = "JMeter Single-Run Load Test Charts"
        chart_ws["A1"].font = Font(bold=True, size=16, color="FFFFFF")
        chart_ws["A1"].fill = PatternFill("solid", fgColor="1F3A5F")
        chart_ws["A2"] = "Charts reflect this uploaded run. Transaction charts show the top 10 where applicable."
        chart_ws.column_dimensions["A"].width = 20
        chart_ws.column_dimensions["J"].width = 4
        row = 4
        for title, image_bytes in charts:
            chart_ws.cell(row=row, column=1, value=title).font = Font(bold=True, size=11, color="1F3A5F")
            img = ExcelImage(BytesIO(image_bytes))
            img.width = 680
            img.height = 360
            chart_ws.add_image(img, f"A{row + 1}")
            row += 22

        for ws in writer.book.worksheets:
            if ws.title == "Charts":
                continue
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="1F3A5F")
                cell.alignment = Alignment(wrap_text=True, vertical="center")
            ws.row_dimensions[1].height = 30
            for col in ws.columns:
                max_len = min(max(len(str(c.value or "")) for c in list(col)[:250]) + 2, 36)
                ws.column_dimensions[get_column_letter(col[0].column)].width = max(12, max_len)
    return output.getvalue()


with st.sidebar:
    st.header("Load test input")
    uploaded = st.file_uploader("Upload one JMeter CSV/JTL file", type=["csv", "jtl"])
    st.caption("Expected core columns: timeStamp, elapsed, label. JMeter CSV output is recommended.")
    st.divider()
    st.subheader("Analysis options")
    top_n = st.slider("Transactions to show in ranked charts", min_value=3, max_value=20, value=10)
    interval = st.selectbox("Time-series interval", ["1 second", "5 seconds", "10 seconds"], index=0)
    show_samples = st.checkbox("Show filterable raw samples", value=True)

if not uploaded:
    st.info("Upload a single JMeter results CSV/JTL file to generate the load-test analysis.")
    st.markdown("""
    **Recommended JMeter result fields:** `timeStamp`, `elapsed`, `label`, `responseCode`,
    `success`, `allThreads`, `Latency`, `Connect`, `bytes`, `sentBytes`.

    **Suggested graphs:** response-time percentiles, response time over time, requests/second,
    active threads, success/failure distribution, status codes, and response-time distribution.
    """)
    st.stop()

try:
    uploaded.seek(0)
    df, warnings = parse_results(uploaded)
except Exception as exc:
    st.error(f"Could not read this results file: {exc}")
    st.stop()

for warning in warnings:
    st.warning(warning)

start_time = df["start_time"].min()
end_time = df["end_time"].max()
duration_s = max((end_time - start_time).total_seconds(), 0.001)
sample_count = len(df)
failure_count = int(df["failed"].sum())
success_count = sample_count - failure_count
error_rate = failure_count / sample_count * 100
overall_throughput = sample_count / duration_s
avg_ms = float(df["elapsed_ms"].mean())
p95_ms = float(df["elapsed_ms"].quantile(0.95))
p99_ms = float(df["elapsed_ms"].quantile(0.99))
max_threads = float(df["threads"].max()) if df["threads"].notna().any() else np.nan

st.subheader("Run overview")
st.caption(
    f"File: **{uploaded.name}** · Start: {start_time.strftime('%Y-%m-%d %H:%M:%S UTC')} · "
    f"Last sample completion: {end_time.strftime('%Y-%m-%d %H:%M:%S UTC')} · "
    f"Estimated duration: {duration_s:.2f} seconds"
)
kpi = st.columns(6)
kpi[0].metric("Total samples", f"{sample_count:,}")
kpi[1].metric("Avg response", f"{avg_ms:,.1f} ms")
kpi[2].metric("P95 response", f"{p95_ms:,.1f} ms")
kpi[3].metric("Overall throughput*", f"{overall_throughput:,.2f} req/s")
kpi[4].metric("Error rate", f"{error_rate:.2f}%")
kpi[5].metric("Max active threads", f"{max_threads:.0f}" if not math.isnan(max_threads) else "N/A")
st.caption("*Throughput is estimated as completed samples divided by the timestamp-derived run span. It may differ from JMeter's official throughput because the CSV may not record exact test start/end boundaries.")

summary = build_transaction_summary(df)
status_summary = (
    df.groupby("response_code", dropna=False)
    .agg(samples=("response_code", "size"), failures=("failed", "sum"))
    .reset_index()
    .sort_values("samples", ascending=False)
)
status_summary["share_pct"] = status_summary["samples"] / sample_count * 100
run_summary = pd.DataFrame([{
    "source_file": uploaded.name,
    "start_time_utc": start_time,
    "last_sample_completion_utc": end_time,
    "estimated_duration_seconds": duration_s,
    "total_samples": sample_count,
    "successful_samples": success_count,
    "failed_samples": failure_count,
    "average_response_ms": avg_ms,
    "median_response_ms": float(df["elapsed_ms"].median()),
    "p90_response_ms": float(df["elapsed_ms"].quantile(0.90)),
    "p95_response_ms": p95_ms,
    "p99_response_ms": p99_ms,
    "min_response_ms": float(df["elapsed_ms"].min()),
    "max_response_ms": float(df["elapsed_ms"].max()),
    "overall_throughput_estimate_req_s": overall_throughput,
    "error_rate_pct": error_rate,
    "max_active_threads": max_threads,
}])

tabs = st.tabs([
    "Overview", "Response Time", "Throughput", "Concurrency", "Errors & Status", "Transactions & Samples", "Export"
])

with tabs[0]:
    st.subheader("Run health indicators")
    c1, c2 = st.columns(2)
    with c1:
        fig = go.Figure(data=[go.Pie(
            labels=["Successful", "Failed"], values=[success_count, failure_count],
            hole=0.58, marker=dict(colors=["#2E8B57", "#C0504D"])
        )])
        fig.update_layout(title="Sample outcome distribution", height=360, margin=dict(t=60,b=20,l=20,r=20))
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        status_plot = status_summary.copy()
        status_plot["response_code"] = status_plot["response_code"].astype(str)
        fig = px.bar(status_plot, x="response_code", y="samples", color="response_code",
                     title="Samples by response code", labels={"response_code":"Response code","samples":"Samples"})
        fig.update_layout(showlegend=False, height=360, margin=dict(t=60,b=20,l=20,r=20))
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Response time by transaction")
    plot_summary = summary.head(top_n).sort_values("average_ms", ascending=True)
    fig = go.Figure()
    fig.add_trace(go.Bar(name="Average", x=plot_summary["average_ms"], y=plot_summary["transaction"], orientation="h"))
    fig.add_trace(go.Bar(name="P95", x=plot_summary["p95_ms"], y=plot_summary["transaction"], orientation="h"))
    fig.update_layout(barmode="group", title=f"Average vs P95 response time (top {min(top_n, len(plot_summary))} by average)",
                      xaxis_title="Response time (ms)", yaxis_title="Transaction", height=max(400, 42*len(plot_summary)+100))
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("Automatic observations")
    observations = []
    slowest = summary.iloc[0]
    observations.append(f"Slowest transaction by average response time: **{slowest['transaction']}** ({slowest['average_ms']:.1f} ms average; P95 {slowest['p95_ms']:.1f} ms).")
    observations.append(f"Overall sample count: **{sample_count:,}** across **{summary.shape[0]}** transaction label(s).")
    if error_rate > 0:
        observations.append(f"**{failure_count} sample(s) failed** ({error_rate:.2f}% error rate). Investigate failures before accepting the run.")
    else:
        observations.append("No failed samples were detected in the uploaded results; this does not by itself prove the application is defect-free.")
    if p95_ms > 2 * max(avg_ms, 0.001):
        observations.append("P95 is more than twice the average response time, suggesting a long-tail latency pattern worth investigating.")
    if df["threads"].notna().any():
        observations.append(f"Maximum active thread count recorded: **{max_threads:.0f}**.")
    for item in observations:
        st.markdown(f"- {item}")

with tabs[1]:
    st.subheader("Response-time analysis")
    percentile_rows = []
    for tx, part in df.groupby("transaction"):
        percentile_rows.append({
            "transaction": tx, "samples": len(part),
            "min_ms": part["elapsed_ms"].min(),
            "median_ms": part["elapsed_ms"].median(),
            "average_ms": part["elapsed_ms"].mean(),
            "p90_ms": part["elapsed_ms"].quantile(0.90),
            "p95_ms": part["elapsed_ms"].quantile(0.95),
            "p99_ms": part["elapsed_ms"].quantile(0.99),
            "max_ms": part["elapsed_ms"].max(),
        })
    percentiles = pd.DataFrame(percentile_rows).sort_values("p95_ms", ascending=False)
    fig = px.bar(percentiles.head(top_n).sort_values("p95_ms"), x="p95_ms", y="transaction", orientation="h",
                 title="P95 response time by transaction", labels={"p95_ms":"P95 response time (ms)","transaction":"Transaction"})
    st.plotly_chart(fig, use_container_width=True)

    fig = px.box(df, x="transaction", y="elapsed_ms", points="outliers",
                 title="Response-time distribution by transaction",
                 labels={"transaction":"Transaction","elapsed_ms":"Response time (ms)"})
    fig.update_layout(xaxis_tickangle=-35, height=520)
    st.plotly_chart(fig, use_container_width=True)

    fig = px.scatter(df, x="relative_seconds", y="elapsed_ms", color="transaction",
                     hover_data=["response_code","success","threads"],
                     title="Individual response times across the run",
                     labels={"relative_seconds":"Seconds since first sample","elapsed_ms":"Response time (ms)"})
    fig.update_layout(height=430)
    st.plotly_chart(fig, use_container_width=True)

    if df["latency_ms"].notna().any():
        fig = px.box(df.dropna(subset=["latency_ms"]), x="transaction", y="latency_ms",
                     title="Server latency by transaction", labels={"latency_ms":"Latency (ms)","transaction":"Transaction"})
        fig.update_layout(xaxis_tickangle=-35, height=430)
        st.plotly_chart(fig, use_container_width=True)
    if df["connect_ms"].notna().any():
        fig = px.histogram(df.dropna(subset=["connect_ms"]), x="connect_ms", nbins=30,
                           title="Connection-time distribution", labels={"connect_ms":"Connect time (ms)"})
        st.plotly_chart(fig, use_container_width=True)
    st.dataframe(percentiles, use_container_width=True, hide_index=True)

with tabs[2]:
    st.subheader("Throughput over time")
    secs = int(interval.split()[0])
    temp = df.copy()
    temp["time_bucket"] = (np.floor(temp["relative_seconds"] / secs) * secs).astype(int)
    throughput_ts = temp.groupby("time_bucket").size().rename("requests").reset_index()
    throughput_ts["requests_per_second"] = throughput_ts["requests"] / secs
    fig = px.line(throughput_ts, x="time_bucket", y="requests_per_second", markers=True,
                  title=f"Requests per second by {secs}-second bucket",
                  labels={"time_bucket":"Seconds since first sample","requests_per_second":"Requests/second"})
    fig.update_layout(height=420)
    st.plotly_chart(fig, use_container_width=True)
    st.caption("The first and last buckets can be partial. The chart divides each bucket's count by the selected bucket width, so edge buckets may understate the rate.")

    st.subheader("Cumulative completed samples")
    cumulative = df.sort_values("start_time").copy()
    cumulative["completed_count"] = np.arange(1, len(cumulative)+1)
    fig = px.line(cumulative, x="relative_seconds", y="completed_count",
                  title="Cumulative sample starts over time",
                  labels={"relative_seconds":"Seconds since first sample","completed_count":"Cumulative samples"})
    st.plotly_chart(fig, use_container_width=True)

with tabs[3]:
    st.subheader("Active threads / virtual users")
    if df["threads"].notna().any():
        threads_ts = df[["relative_seconds","threads"]].dropna().sort_values("relative_seconds")
        fig = px.line(threads_ts, x="relative_seconds", y="threads", markers=True,
                      title="Active threads recorded with samples",
                      labels={"relative_seconds":"Seconds since first sample","threads":"Active threads"})
        st.plotly_chart(fig, use_container_width=True)
        st.caption("JMeter's allThreads value is sampled at result events, not a continuous thread timeline. Use it as an approximation of concurrency during the run.")
        st.metric("Maximum recorded active threads", f"{max_threads:.0f}")
    else:
        st.info("This result file does not include allThreads. Enable the field in JMeter's CSV/JTL output to chart concurrency.")

with tabs[4]:
    st.subheader("Failures and response codes")
    left, right = st.columns(2)
    with left:
        fig = px.pie(status_summary, names="response_code", values="samples", hole=0.45,
                     title="Response-code mix")
        st.plotly_chart(fig, use_container_width=True)
    with right:
        fail_by_tx = df.groupby("transaction").agg(
            samples=("transaction","size"), failures=("failed","sum")
        ).reset_index()
        fail_by_tx["error_rate_pct"] = 100 * fail_by_tx["failures"] / fail_by_tx["samples"]
        fig = px.bar(fail_by_tx.sort_values("error_rate_pct", ascending=False).head(top_n),
                     x="transaction", y="error_rate_pct", color="error_rate_pct",
                     title="Error rate by transaction",
                     labels={"transaction":"Transaction","error_rate_pct":"Error rate (%)"})
        fig.update_layout(xaxis_tickangle=-35)
        st.plotly_chart(fig, use_container_width=True)
    failed_df = df[df["failed"]].copy()
    st.markdown("**Failed samples**")
    if failed_df.empty:
        st.success("No failed samples were identified by the success flag / response-code logic.")
    else:
        st.dataframe(failed_df[["start_time","transaction","elapsed_ms","response_code","url","threads"]], use_container_width=True, hide_index=True)
    st.dataframe(status_summary, use_container_width=True, hide_index=True)

with tabs[5]:
    st.subheader("Transaction summary")
    st.dataframe(summary, use_container_width=True, hide_index=True)
    if show_samples:
        st.subheader("Filterable raw samples")
        transaction_options = ["All"] + sorted(df["transaction"].unique().tolist())
        selected_tx = st.selectbox("Filter by transaction", transaction_options)
        result_options = ["All", "Success", "Failure"]
        selected_result = st.selectbox("Filter by result", result_options)
        filtered = df.copy()
        if selected_tx != "All":
            filtered = filtered[filtered["transaction"] == selected_tx]
        if selected_result != "All":
            filtered = filtered[filtered["result"] == selected_result]
        st.dataframe(filtered, use_container_width=True, hide_index=True)

with tabs[6]:
    st.subheader("Download analysis")
    st.write("The Excel workbook includes the run summary, transaction-level statistics, response-code counts, and every parsed sample.")
    excel_bytes = make_excel(df, summary, status_summary, run_summary)
    st.download_button(
        "📥 Download Excel analysis",
        data=excel_bytes,
        file_name="jmeter_single_run_analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.download_button(
        "📥 Download normalized samples CSV",
        data=df.to_csv(index=False).encode("utf-8"),
        file_name="jmeter_normalized_samples.csv",
        mime="text/csv",
    )
    st.subheader("Run summary")
    st.dataframe(run_summary, use_container_width=True, hide_index=True)

st.divider()
st.caption("Interpret results in context: workload, environment, test duration, and sample count affect conclusions. Use only systems you own or are authorized to test.")
