# Delivery Monitoring with EricMon

![Delivery Monitoring](delivery_monitoring.png)

## Overview

This Python script, `delivery_check.py`, is designed for monitoring daily deliveries through the Outlook Sent Items folder. It retrieves delivery information from an EricMon database and matches it with emails in the Sent Items folder to track the status of deliveries. The script generates an HTML report and sends it via email to relevant stakeholders.

## Features

- Retrieve delivery information from an EricMon database.
- Filter Sent Items in Outlook based on today's date.
- Match email subjects with delivery records and track delivery status.
- Generate an HTML report with a table displaying email subjects, times, expected times, statuses, and owners.
- Highlight late deliveries and unsuccessful deliveries in the report.
- Send the report via email using Outlook.

## Prerequisites

Before using this script, ensure you have the following prerequisites:

- Python 3.x installed on your system.
- The `pywin32` library to work with Windows-specific features.
- Access to an Outlook email account.
- Access to an EricMon database.
- Proper configuration of the Outlook email account and database credentials in the script.

## Usage

1. Clone this repository to your local machine.

2. Install the required dependencies using pip:

```
pip install -r requirements.txt
```
## Run the script:
```
python delivery_check.py
```

## CSS Styling
The HTML report is styled using the EricMon.css file, which defines the appearance of the report's table, headers, and row colors. You can customize the styling by modifying this CSS file.

## Conclusion
The Delivery Monitoring with EricMon script simplifies the management of delivery tracking by automating the process of matching delivery records with sent emails. It ensures that stakeholders are informed of delivery statuses in real-time, enhancing efficiency and reducing manual tracking efforts.