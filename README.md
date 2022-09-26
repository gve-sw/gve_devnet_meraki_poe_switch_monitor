# Meraki PoE Switch Monitor
This prototype checks the PoE usage of switches in a Meraki organization. The code will create an Excel workbook that contains 6 spreadsheets: Low Power Switches, High Power Switches, No Power Switches, Switchport Power Usage, Disconnected Ports, and Offline Switches. The Low Power Switches, High Power Switches, and No Power Switches spreadsheets will contain information about switches that use 67W or less of PoE, more than 67W of PoE, and 0W of PoE, respectively. The Switchport Power Usage spreadsheet contains information about all connected switchports in the organization, including their individual PoE usage. The Disconnected Ports lists out the switchports in the organization that are not connected. Lastly, the Offline Switches spreadsheet will list out the switches that are either offline, dormant, or alerting.

## Contacts
* Danielle Stacy
*  Jordan Coaten

## Solution Components
* Meraki
* Python 3.9
* Pandas

## Prerequisites
#### Meraki API Keys
In order to use the Meraki API, you need to enable the API for your organization first. After enabling API access, you can generate an API key. Follow these instructions to enable API access and generate an API key:
1. Login to the Meraki dashboard
2. In the left-hand menu, navigate to `Organization > Settings > Dashboard API access`
3. Click on `Enable access to the Cisco Meraki Dashboard API`
4. Go to `My Profile > API access`
5. Under API access, click on `Generate API key`
6. Save the API key in a safe place. The API key will only be shown once for security purposes, so it is very important to take note of the key then. In case you lose the key, then you have to revoke the key and a generate a new key. Moreover, there is a limit of only two API keys per profile.

> For more information on how to generate an API key, please click [here](https://developer.cisco.com/meraki/api-v1/#!authorization/authorization). 

> Note: You can add your account as Full Organization Admin to your organizations by following the instructions [here](https://documentation.meraki.com/General_Administration/Managing_Dashboard_Access/Managing_Dashboard_Administrators_and_Permissions).

## Installation/Configuration
1. Clone this repository with `git clone https://github.com/gve-sw/gve_devnet_meraki_poe_switch_monitor`
2. Add Meraki API key that you retrieved in the Prerequisites section and the name of your Meraki Organization to environment variables
```python
TOKEN = "API key goes here"
ORG_NAME = "organization name goes here"
```
3. Set up a Python virtual environment. Make sure Python 3 is installed in your environment, and if not, you may download Python [here](https://www.python.org/downloads/). Once Python 3 is installed in your environment, you can activate the virtual environment with the instructions found [here](https://docs.python.org/3/tutorial/venv.html).
4. Install the requirements with `pip3 install -r requirements.txt`

## Usage
To run the program, use the command:
```
$ python3 app.py
```

Once the code is finished running, an Excel workbook called poe_switches.xlsx will be created in the directory. This workbook will contain six different spreadsheets: Low Power Switches, High Power Switches, No Power Switches, Switchport Power Usage, Disconnected Switchports, and Offline Switches.

# Screenshots

![/IMAGES/0image.png](/IMAGES/0image.png)


Output from the code
![/IMAGES/poe_switch_monitor_output.png](/IMAGES/poe_switch_monitor_output.png)

Excel workbook
![/IMAGES/poe_switches_screenshot.png](/IMAGES/poe_switches_screenshot.png)

### LICENSE

Provided under Cisco Sample Code License, for details see [LICENSE](LICENSE.md)

### CODE_OF_CONDUCT

Our code of conduct is available [here](CODE_OF_CONDUCT.md)

### CONTRIBUTING

See our contributing guidelines [here](CONTRIBUTING.md)

#### DISCLAIMER:
<b>Please note:</b> This script is meant for demo purposes only. All tools/ scripts in this repo are released for use "AS IS" without any warranties of any kind, including, but not limited to their installation, use, or performance. Any use of these scripts and tools is at your own risk. There is no guarantee that they have been through thorough testing in a comparable environment and we are not responsible for any damage or data loss incurred with their use.
You are responsible for reviewing and testing any scripts you run thoroughly before use in any non-testing environment.
