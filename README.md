# Outlook Delivery Monitoring System

The **Outlook Delivery Monitoring System** is a Python-based tool designed to track and monitor daily deliveries through Microsoft Outlook's Sent Items folder. This system is particularly useful for businesses and organizations that rely on email communication for critical deliveries and need to ensure the timely delivery of important information.

## Table of Contents
- [Overview](#overview)
- [Key Features](#key-features)
- [Installation](#installation)
- [Usage](#usage)
- [License](#license)
- [Acknowledgments](#acknowledgments)
- [Contributing](#contributing)

## Overview

Email communication is a fundamental part of modern business operations. Ensuring that important messages, reports, and notifications are delivered promptly is crucial. The **Outlook Delivery Monitoring System** automates the process of monitoring deliveries and provides insights into the status of each delivery. It accomplishes this by comparing sent emails with expected deliveries and categorizing them based on their status.

## Key Features

- **Automated Monitoring:** The system automatically scans the Sent Items folder in Microsoft Outlook to identify sent emails.

- **Status Categorization:** It categorizes emails based on their delivery status, including "Success," "Late Delivery," "Pending Delivery," and "Delivery Unsuccessful."

- **Real-time Reporting:** Users receive real-time reports summarizing the delivery status of sent emails.

- **Customizable Alerts:** The system can be configured to send alerts or notifications when certain conditions are met, such as late deliveries or delivery failures.

- **Email Owners:** It associates each email with its respective owner or sender, providing insights into responsible parties.

## Installation

To get started with the **Outlook Delivery Monitoring System**, follow these steps:

1. Clone this repository to your local machine.

2. Install the required dependencies using the following command:

```
pip install -r requirements.txt
```

## Run the project with:
```
python delivery_check.py
```
## Usage
To extract and monitor deliveries from a specific time frame, the system uses the Sent Items folder in Microsoft Outlook. It compares the sent emails with a list of expected deliveries and categorizes them based on their delivery status.
