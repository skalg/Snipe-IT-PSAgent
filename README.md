# Snipe-IT Asset Management Script

## Overview

This PowerShell script is designed as an agent for the Snipe-IT asset management system. It performs the following tasks:

1. **Retrieve Computer Information**: Gathers detailed information about the computer, such as model, serial number, RAM, CPU, OS version, kernel, storage type, storage capacity, MAC addresses, IP addresses, and more.
2. **Search and Create Models**: Searches for the computer model in Snipe-IT. If it doesn't exist, the script creates a new model.
3. **Search and Create Assets**: Searches for the computer asset in Snipe-IT using the serial number. If it doesn't exist, the script creates a new asset and populates various custom fields with the gathered information.
4. **Update Existing Assets**: Updates the existing asset's name and custom fields if there are any changes.
5. **Hyper-V VM Information**: Checks if Hyper-V is installed and, if so, lists the VMs on the host and updates a custom field with this information.

## Prerequisites

- **Snipe-IT API Token**: You need a valid Snipe-IT API token with appropriate permissions to access and modify models and assets.
- **PowerShell**: The script is written in PowerShell and requires PowerShell to execute.

## Setup

1. Clone the repository or download the script to your local machine.
2. Open the script in a text editor and configure the following variables:
    - `$SnipeItApiUrl`: The base URL of your Snipe-IT instance.
    - `$SnipeItApiToken`: Your Snipe-IT API token.
    - `$category_id`: The category ID for your models (default: 3 for desktops, 2 for laptops).
    - `$fieldset_id`: The fieldset ID for your models.
3. Update the custom fields in the `Get-CustomFields` function to match the ones in your Snipe-IT installation.

## Usage

Run the script using PowerShell:

```powershell
.\snipeit-asset-management.ps1
