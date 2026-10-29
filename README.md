# JMeter Single-Run Load Test Analysis Dashboard

## Project Overview

This project is a Python-based performance test analysis dashboard built with Streamlit. It analyses Apache JMeter load-test results and presents performance metrics, interactive visualisations, transaction-level summaries, and downloadable reports.

The project demonstrates practical skills in performance testing, Python data analysis, test-result interpretation, and reporting.

## Key Features

- **Test Run Overview:** Total samples, average response time, P95 response time, throughput estimate, error rate, and maximum recorded active threads.
- **Response Time Analysis:** Average and P95 response times by transaction, response-time distribution, and individual sample timings.
- **Throughput Analysis:** Requests over time and cumulative completed samples.
- **Concurrency Analysis:** Active-thread analysis when thread data is available.
- **Error Analysis:** Failed samples, HTTP response codes, and transaction-level error rates.
- **Additional Metrics:** Server latency and connection-time analysis when the required fields are available.
- **Excel Reporting:** Downloadable workbook containing summary tables, sample data, and embedded charts.
- **CSV Export:** Downloadable normalized sample data for further analysis.

## Technologies Used

- Python
- Streamlit
- Pandas
- NumPy
- Plotly
- Matplotlib
- OpenPyXL
- Apache JMeter

## Project Files

- `jmeter_single_run_load_test_analysis.py` — Main dashboard application.
- `requirements_single_run_analysis.txt` — Required Python packages.
- `jmeter_single_run_analysis.xlsx` — Optional sample analysis workbook.
- `jmeter_normalized_samples.csv` — Optional sample normalized dataset.

## Installation

### 1. Install Python

Install a compatible version of Python and ensure Python is available from your terminal.

### 2. Install dependencies

Open a terminal in the project folder and run:

```bash
python -m pip install -r requirements_single_run_analysis.txt
```

### 3. Run the dashboard

```bash
python -m streamlit run jmeter_single_run_load_test_analysis.py
```

### 4. Analyse JMeter results

Open the local Streamlit URL displayed in the terminal, upload a JMeter CSV/JTL results file, explore the dashboard, and download the generated reports.

## Performance Metrics Explained

- **Average Response Time:** Mean response time across recorded samples.
- **P95 Response Time:** The 95th percentile response time; 95% of recorded responses are at or below this value.
- **Throughput Estimate:** An estimate of completed samples per unit of time based on the timestamps available in the input data.
- **Error Rate:** Percentage of recorded samples classified as failed.
- **Active Threads:** Recorded concurrent JMeter threads, when available in the results file.

## Important Limitations

- Metrics depend on the fields available in the input CSV/JTL file.
- The calculated throughput estimate may differ from JMeter's official throughput because the exact test boundaries may not be present in the results.
- Active-thread values represent recorded result events, not necessarily continuous concurrency measurements.
- Performance conclusions should consider the test environment, workload, ramp-up, duration, and application behaviour.

## Future Enhancements

- Automated email delivery of performance reports.
- Baseline-versus-current test comparison.
- Automated performance regression detection.
- Additional report formats and analysis features.

Developed as a portfolio project to demonstrate performance testing, Python-based analytics, and automated reporting.
