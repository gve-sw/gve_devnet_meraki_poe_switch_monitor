""" Copyright (c) 2020 Cisco and/or its affiliates.
This software is licensed to you under the terms of the Cisco Sample
Code License, Version 1.1 (the "License"). You may obtain a copy of the
License at
           https://developer.cisco.com/docs/licenses
All use of the material herein must be in accordance with the terms of
the License. All rights not expressly granted by the License are
reserved. Unless required by applicable law or agreed to separately in
writing, software distributed under the License is distributed on an "AS
IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
or implied.
"""
import meraki
import os
import pandas as pd

from dotenv import load_dotenv

#get environmental variables
load_dotenv()

API_KEY = os.environ['TOKEN']
ORG_NAME = os.environ['ORG_NAME']
dashboard = meraki.DashboardAPI(API_KEY, suppress_logging=True)

def main():
    #get org id
    org_id = get_org_id()
    if org_id is None:
        #given organization name did match an organization associated with Meraki API key
        print("There was an issue getting the org id. Check that the organization name is correct.")
    else:
        print("Retrieved organization ID")
        #get networks from organization
        networks = get_networks(org_id)
        print("Retrieved networks from the organization")
        #need data structure that maps network names to their ids
        net_name_to_id = create_network_dict(networks)
        #get all the switches from the organization)
        switches = get_switches(org_id)
        print("Retrieved switches from the organization")
        #pull relevant information from switches and add it to dictionary
        switch_info = get_switch_info(switches, net_name_to_id)
        #get port information from each of the switches
        switch_port_statuses = get_device_details(switches)
        print("Retrieved port statuses from the switches")
        #get the availability of the switches - online or offline or dormant
        switch_availabilities = get_switch_availabilities(org_id)
        print("Retrieved availabilities of switches")
        #get the online and offline switches from the switches
        parsed_switches = parse_online_switches(switch_availabilities, switch_info)
        #parse out the port information from the switches
        spreadsheet_data = parse_port_statuses(switch_availabilities, switch_port_statuses, parsed_switches)
        print("Starting to create Excel workbook")
        #create Excel workbook with the low power, high power, and no power switches
        create_excel_workbook(spreadsheet_data)
        print("Created Excel workbook")

'''This function returns the organization id that is associated
with the organization name given in the environmental variables
in the .env file'''
def get_org_id():
    response = dashboard.organizations.getOrganizations()
    for org in response:
        if org['name'] == ORG_NAME:
            org_id = org['id']
            return org_id

    #if no organization is found with the name then None is returned
    return None

'''This function returns the networks that are associated with
the given organization id'''
def get_networks(org_id):
    networks = dashboard.organizations.getOrganizationNetworks(
        org_id, total_pages='all')
    return networks

'''This function creates a dictionary that has network ids as
its keys and the names of the networks as values. This data
will be referenced later in the code when we are given the
network ids but need to reference the network names'''
def create_network_dict(networks):
    net_name_to_id = {}
    for network in networks:
        network_id = network['id']
        network_name = network['name']

        net_name_to_id[network_id] = network_name

    return net_name_to_id

'''This function retrieves all the switches in the
organization. The function then returns the switches in a
list.'''
def get_switches(org_id):
    switches = dashboard.organizations.getOrganizationDevices(org_id, total_pages='all', productTypes=['switch'])

    return switches

'''This function pulls the necessary information from the
switches list about each switch that will be used in the
spreadsheet. It then compiles this information into a
dictionary with the switch's serial number as the key and
dictionary that specifies the name of the network the
switch is in, the name of the switch, and the model of the
switch.'''
def get_switch_info(switches, net_name_to_id):
    switch_info = {}
    for switch in switches:
        serial = switch['serial']
        network_id = switch['networkId']
        net_name = net_name_to_id[network_id]
        model = switch['model']
        if 'name' in switch.keys():
            name = switch['name']
        else:
            name = '' #in some cases, the devices are not assigned a name

        switch_info[serial] = {
            "network": net_name,
            "name": name,
            "model": model
        }

    return switch_info

'''This function retrieves the switchport information from
each switch. The port information will give PoE usage for
each port of the switch over a certain time period. The
function returns a dictionary that has the switch serial
number as the key and a list of the port statuses as the
value.'''
def get_device_details(switches):
    switch_statuses = {}
    for switch in switches:
        serial = switch['serial']
        port_statuses = dashboard.switch.getDeviceSwitchPortsStatuses(
            serial, timespan=3600) #specify the timespan as 1hr to aggregate the information in the last hour
        switch_statuses[serial] = port_statuses

    return switch_statuses

'''This function gets the availabilities of the switches
in the organization. This information will be return as a
dictionary with the switch serial as they key and the
availability of the switch (online, offline, dormant,
alerting) as the value.'''
def get_switch_availabilities(org_id):
    switch_availabilities = {}
    availabilities = dashboard.organizations.getOrganizationDevicesAvailabilities(
        org_id, productTypes=['switch'])

    for availability in availabilities:
        serial = availability['serial']
        status = availability['status']

        switch_availabilities[serial] = status

    return switch_availabilities

'''This function parses through the switch availabilities and
structures the information to be added to the spreadsheet. The
function returns a dictionary with keys 'online' and 'offline'
and the values are lists of the online and offline switches
respectively. The lists consist of dictionaries that hold the
serial number, name, model, network, and availability of the
switch.'''
def parse_online_switches(switch_availabilities, switch_info):
    onlineSwitches = []
    offlineSwitches = []

    for switch in switch_info:
        power_total = 0
        network = switch_info[switch]['network']
        switch_name = switch_info[switch]['name']
        model = switch_info[switch]['model']
        switch_status = switch_availabilities[switch]

        switch_dict = {
            'serial': switch,
            'name': switch_name,
            'model': model,
            'network': network,
            'status': switch_status
        }

        if switch_status != 'online':
            offlineSwitches.append(switch_dict)
        else:
            onlineSwitches.append(switch_dict)

    parsed_switches = {
        'online': onlineSwitches,
        'offline': offlineSwitches
    }

    return parsed_switches

'''This function parses through the port statuses and
structures the information to be added to the spreadsheet.
The function returns a dictionary with the switches using
no PoE, low PoE (<= 67W), and high PoE(> 67W), as well as
the disconnected ports.'''
def parse_port_statuses(switch_availabilities, switch_port_statuses, parsed_switches):
    noPoE = []
    lowPoE = []
    highPoE = []
    portPoE = []
    disconnectedPorts = []

    online_switches = parsed_switches['online']
    offline_switches = parsed_switches['offline']

    for switch in online_switches:
        serial = switch['serial']
        power_total = 0

        for port_status in switch_port_statuses[serial]:
            if 'powerUsageInWh' in port_status.keys():
                power_usage = port_status['powerUsageInWh']
                port_id = port_status['portId']
                enabled = port_status['enabled']
                port_status = port_status['status']

                switch_port = {
                    'portId': port_id,
                    'switchSerial': serial,
                    'enabled': enabled,
                    'portStatus': port_status,
                    'powerUsage': power_usage
                }

                switch_status = switch_availabilities[serial]

                if port_status == 'Disconnected':
                    disconnectedPorts.append(switch_port)
                elif port_status == 'Connected':
                    portPoE.append(switch_port)

                    if switch_status == 'online':
                        power_total += power_usage

        switch_dict = switch
        switch_dict['powerUsageOfSwitch'] = power_total

        if power_total == 0:
            noPoE.append(switch_dict)
        elif power_total <= 67:
            lowPoE.append(switch_dict)
        elif power_total > 67:
            highPoE.append(switch_dict)

    spreadsheet_lists = {
        "noPoE": noPoE,
        "lowPoE": lowPoE,
        "highPoE": highPoE,
        "disconnectedPorts": disconnectedPorts,
        "portPoE": portPoE,
        "offlineSwitches": offline_switches
    }

    return spreadsheet_lists

'''This function creates an Excel workbook with six
different worksheets. The worksheets reflect information
about the switches that are currently using no PoE,
those that are using low PoE (<= 67W), those that
are using high PoE (> 67W), information about all the ports
using PoE, information about all the disconnected ports,
and information about offline switches.'''
def create_excel_workbook(spreadsheet_lists):
    noPoE = spreadsheet_lists['noPoE']
    lowPoE = spreadsheet_lists['lowPoE']
    highPoE = spreadsheet_lists['highPoE']
    disconnectedPorts = spreadsheet_lists['disconnectedPorts']
    portPoE = spreadsheet_lists['portPoE']
    offlineSwitches = spreadsheet_lists['offlineSwitches']

    with pd.ExcelWriter('poe_switches.xlsx') as writer:
        lowPoE_df = pd.DataFrame.from_dict(lowPoE)
        highPoE_df = pd.DataFrame.from_dict(highPoE)
        noPoE_df = pd.DataFrame.from_dict(noPoE)
        portPoE_df = pd.DataFrame.from_dict(portPoE)
        disconnectedPort_df = pd.DataFrame.from_dict(disconnectedPorts)
        offlineSwitches_df = pd.DataFrame.from_dict(offlineSwitches)

        lowPoE_df.to_excel(writer, sheet_name='Low Power Switches')
        highPoE_df.to_excel(writer, sheet_name='High Power Switches')
        noPoE_df.to_excel(writer, sheet_name='No Power Switches')
        portPoE_df.to_excel(writer, sheet_name='Switchport Power Usage')
        disconnectedPort_df.to_excel(writer, sheet_name='Disconnected Ports')
        offlineSwitches_df.to_excel(writer, sheet_name='Offline Switches')

if __name__ == "__main__":
    main()
