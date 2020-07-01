#!/usr/bin/python

import sys
import pandas as pd
import re
from datetime import datetime
from getpass import getpass
from urllib3 import disable_warnings
from cobra.mit.access import MoDirectory
from cobra.mit.session import LoginSession
from cobra.mit.request import ConfigRequest
from cobra.model.fv import AEPg, Ap, BD, Tenant, RsPathAtt
from os import name, path, system, remove
from requests.exceptions import HTTPError, ConnectionError
from termcolor import colored
from time import sleep

disable_warnings()
pd.set_option('mode.chained_assignment', None)
excel_file = "EPGS_WITHOUT_BINDING.xlsx"
epg_regex = r'(uni/tn-(.+)/ap-(.+)/epg-(.+))'
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
filename = "Static_Ports_Deployment_" + timestamp + ".log"
count = 0


def clear():
    """Clear function to clean the CLI screen.
    Does not return any value."""
    # for windows
    if name == "nt":
        _ = system("cls")
    # for mac and linux(here, os.name is 'posix')
    else:
        _ = system("clear")


def banner():
    color1, color2, color3, color4, color5 = "blue", "green", "cyan", "magenta", "yellow"
    cisco_colors = [
        colored('.aMMMb',  color1), colored('dMP',  color2), colored('.dMMMb',  color3),
        colored('.aMMMb',  color4), colored('.aMMMb',  color5), colored('dMP"VMP',  color1),
        colored('amr',  color2), colored('dMP"',  color3), colored('VP',  color3), colored('dMP"VMP',  color4),
        colored('dMP"dMP',  color5), colored('dMP',  color1), colored('dMP',  color2), colored('VMMMb',  color3),
        colored('dMP',  color4), colored('dMP',  color5), colored('dMP',  color5), colored('dMP.aMP',  color1),
        colored('dMP',  color2), colored('dP',  color3), colored('.dMP',  color3), colored('dMP.aMP',  color4),
        colored('dMP.aMP',  color5), colored('VMMMP"',  color1), colored('dMP',  color2), colored('VMMMP"',  color3),
        colored('VMMMP"',  color4), colored('VMMMP"',  color5)
    ]
    space = " " * 36
    cisco = f"{space}" + "   %s  %s %s  %s  %s \n" \
            f"{space}" + "  %s %s %s %s %s %s \n" \
            f"{space}" + " %s     %s  %s  %s     %s %s  \n" \
            f"{space}" + "%s %s %s %s %s %s   \n" \
            f"{space}" + "%s %s  %s  %s  %s    "
    clear()
    print()
    print(cisco % tuple(cisco_colors))


def thank_you():
    color1, color2, color3, color4 = "blue", "green", "yellow", "magenta",
    color5, color6, color7 = "white", "yellow", "red"
    color8 = color2
    thanks_colors = [
        colored("dMMMMMMP", color=color1), colored('dMP', color=color2), colored('dMP', color=color2),
        colored('.aMMMb', color=color3), colored('dMMMMb', color=color4), colored('dMP', color=color5),
        colored('dMP', color=color5), colored('dMP', color=color6), colored('dMP', color=color6),
        colored('.aMMMb', color=color7), colored('dMP', color=color8), colored('dMP', color=color8),
        colored('dMP', color=color1), colored('dMP', color=color2), colored('dMP', color=color2),
        colored('dMP"dMP', color=color3), colored('dMP', color=color4), colored('dMP', color=color4),
        colored('dMP.dMP', color=color5), colored('dMP.dMP', color=color6), colored('dMP"dMP', color=color7),
        colored('dMP', color=color8), colored('dMP', color=color8), colored('dMP', color=color1),
        colored('dMMMMMP', color=color2), colored('dMMMMMP', color=color3), colored('dMP', color=color4),
        colored('dMP', color=color4), colored('dMMMMK', color=color5), colored('VMMMMP', color=color6),
        colored('dMP', color=color7), colored('dMP', color=color7), colored('dMP', color=color8),
        colored('dMP', color=color8), colored('dMP', color=color1), colored('dMP', color=color2),
        colored('dMP', color=color2), colored('dMP', color=color3), colored('dMP', color=color3),
        colored('dMP', color=color4), colored('dMP', color=color4), colored('dMP"AMF', color=color5),
        colored('dA', color=color6), colored('.dMP', color=color6), colored('dMP.aMP', color=color7),
        colored('dMP.aMP', color=color8), colored('dMP', color=color1), colored('dMP', color=color2),
        colored('dMP', color=color2), colored('dMP', color=color3), colored('dMP', color=color3),
        colored('dMP', color=color4), colored('dMP', color=color4), colored('dMP', color=color5),
        colored('dMP', color=color5), colored('VMMMP"', color=color6), colored('VMMMP"', color=color7),
        colored('VMMMP"', color=color8)
    ]
    space = " " * 20
    thanks = f"{space}" + " %s %s %s %s  %s  %s %s        %s %s %s  %s %s \n" \
             f"{space}" + "   %s   %s %s %s %s %s %s        %s %s %s %s  \n" \
             f"{space}" + "  %s   %s %s %s %s %s          %s %s %s %s %s   \n" \
             f"{space}" + " %s   %s %s %s %s %s %s %s        %s %s %s %s    \n" \
             f"{space}" + "%s   %s %s %s %s %s %s %s %s         %s  %s  %s     \n"
    clear()
    print()
    print(thanks % tuple(thanks_colors))
    sleep(4)


def welcome():
    """Welcome message to provide information about current and future functionality.

    """
    print("%s %s.\n" % ("""
    This script will provide a series of options to choose from, to deploy Static Path Bindings in an ACI fabric.
    The options provided are meant to be increased in future versions of this script.
    If you find any issues with the script, please copy the error message and send it to""",
                      colored("miguecor@cisco.com", color="blue")))
    print("%s%s\n" % ((colored("    NOTE: ", color="magenta")),
                           colored("This script will look better if you adjust your terminal "
                                   "width to a value of 120 columns", color="yellow")))
    sleep(2)
    input(colored(" Please hit 'Enter' when you are ready to start... ", color="magenta"))
    clear()


def get_apic_info():
    """Input function to gather APIC IP/FQDN and authentication parameters.
    Returns strings with URL, USR & PWD.
    """
    print(colored("%s", "blue") % "\n APIC AUTHENTICATION INFO... \n")
    url = ("https://" + input("%s".rjust(len("%s") + 4) % "Enter the APIC's IP/Hostname/FQDN: "))
    usr = input("%s".rjust(len("%s") + 4) % "Enter your username: ")
    pwd = getpass("%s".rjust(len("%s") + 4) % "Enter your password: ")
    return url, usr, pwd


def login(login_info):
    try:
        session = LoginSession(*login_info)
        mo_dir = MoDirectory(session)
        mo_dir.login()
        print("%s" % colored("\n Authentication successful!", color="green", attrs=["bold"]))
        sleep(2)
        clear()
        return mo_dir

    except HTTPError as error:
        print("There's an error in you APIC's username and/or password.\n",
              "Please verify and try again.\n",
              colored("Error: ", color="red"), colored(error, "yellow"))
        print()
        sys.exit()

    except ConnectionError as error:
        print("The provided APIC's IP Address is either unresponsive or wrong.\n",
              "Please verify and try again.\n",
              colored("Error: ", color="red"), colored(error, "yellow"))
        print()
        sys.exit()


def select(rng):
    retries = 5
    for i in range(retries):
        try:
            selection = int(input(colored("%s", color="magenta") % " Enter your selection -->  "))
            while selection not in range(1, rng + 1):
                print("%s" % colored(f"\n {selection} is not a valid option. Please try again.\n", color="red"))
                selection = int(input(colored("%s", color="magenta") % " Enter your selection -->  "))
            return selection

        except ValueError:
            print(colored("%s" % "\n 'Empty' is not a valid selection.  "
                                 "Please use a valid option from the provided list.", color="red"))
            print(colored(" You can try %s more time(s).\n", "yellow") % ((retries - 1) - i))
            continue
    goodbye()


def choose_binding_type():
    print(colored("%s", "blue") % "\n What type of Static Path Binding do you want to configure?\n\n",
          "%s".rjust(len("%s") + 4) % "1) Single Interface\n",
          "%s".rjust(len("%s") + 4) % "2) Regular Port-Channel\n",
          "%s".rjust(len("%s") + 4) % "3) Virtual Port-Channel\n"
          )
    selection = select(3)
    if selection == 1:
        print("\n Interface selection: ", colored("%s", "yellow") % "SINGLE INTERFACE\n")
        pod = input("%s".rjust(len("%s") + 4) % colored("Enter the Pod-ID.  Only numeric values (e.g. 1, 2, 3): ",
                                                        color="magenta"))
        leaf = input("%s".rjust(len("%s") + 4) % colored("Enter the switch node-ID.  Only numeric values "
                                                         "(e.g. 101, 102, 201, 202): ", color="magenta"))
        access = input("%s".rjust(len("%s") + 4) % colored("Enter the interface-ID in the form 'module/port' "
                                                           "(e.g. 1/1, 1/23, 2/16): ", color="magenta"))
        bind_tdn = "topology/pod-%s/paths-%s/pathep-[eth%s]" % (pod, leaf, access)
    elif selection == 2:
        print("\n Interface selection: ", colored("%s", "yellow") % "REGULAR PORT-CHANNEL\n")
        pod = input("%s".rjust(len("%s") + 4) % colored("Enter the Pod-ID.  Only numeric values (e.g. 1, 2, 3): ",
                                                        color="magenta"))
        leaf = input("%s".rjust(len("%s") + 4) % colored("Enter the switch node-ID.  Only numeric values "
                                                         "(e.g. 101, 102, 201, 202): ", color="magenta"))
        portchannel = input("%s".rjust(len("%s") + 4) % colored("Enter the Port-Channel Interface Policy Group name "
                                                                "(e.g. server101, esx2_ifPolGrp): ", color="magenta"))
        bind_tdn = "topology/pod-%s/paths-%s/pathep-[%s]" % (pod, leaf, portchannel)
    else:
        print("\n Interface selection: ", colored("%s", "yellow") % "VIRTUAL PORT-CHANNEL\n")
        pod = input("%s".rjust(len("%s") + 4) % colored("Enter the Pod-ID.  Only numeric values (e.g. 1, 2, 3): ",
                                                        color="magenta"))
        leaf1 = input("%s".rjust(len("%s") + 4) % colored("Enter the switch 1 node-ID.  Only numeric values "
                                                          "(e.g. 101, 102, 201, 202): ", color="magenta"))
        leaf2 = input("%s".rjust(len("%s") + 4) % colored("Enter the switch 2 node-ID.  Only numeric values "
                                                          "(e.g. 101, 102, 201, 202): ", color="magenta"))
        vpc = input("%s".rjust(len("%s") + 4) % colored("Enter the vPC Interface Policy Group name "
                                                        "(e.g. server101, esx2_ifPolGrp): ", color="magenta"))
        bind_tdn = "topology/pod-%s/protpaths-%s-%s/pathep-[%s]" % (pod, leaf1, leaf2, vpc)
    clear()
    return bind_tdn


def where_to_deploy(mo_dir: MoDirectory) -> tuple:
    """Takes the moDir object from the ACI fabric and dissects it into two DataFrames.

    After choosing from a number of options to select the way to deploy the Static Path Bindings, the information
    gathered from the ACI fabric will be returned in a tuple containing two pd.DataFrame objects.

    :param mo_dir: The MoDirectory object obtained after successfully login into the ACI fabric.
    :type mo_dir: MoDirectory
    :return Tuple of returned :class:`pd.DataFrame` objects.
    :rtype: tuple
    """
    available_color = "green"
    print(colored("%s", "blue") % "\n On what EPGs do you want to deploy the Static Path Binding(s)?\n\n",
          "%s".rjust(len("%s") + 4) % colored("1) Deploy on all EPGs from an AppProfile\n",
                                              color=f"{available_color}", attrs=["bold"]),
          "%s".rjust(len("%s") + 4) % "2) Deploy on all EPGs from a VRF\n",
          "%s".rjust(len("%s") + 4) % "3) Deploy on same EPGs as another interface/port-channel/vPC binding\n",
          "%s".rjust(len("%s") + 4) % "4) Deploy from a CSV file\n",
          "%s".rjust(len("%s") + 4) % "5) Deploy on a single EPG\n"
          "\n%s %s %s %s\n" % (colored(" IMPORTANT:", "red", attrs=["blink", "bold"]),
                               colored("Only the options in"),
                               colored(f"{available_color}".upper(), color=f"{available_color}", attrs=["bold"]),
                               colored("are available at the moment.")
                               )
          )
    selection = select(1)
    if selection == 1:
        epg_df, bind_df = epgs_from_app_p(mo_dir)
        return epg_df, bind_df


def epgs_from_app_p(mo_dir):
    print("\n Deployment selection: ", colored("%s", "yellow") % "ALL EPGs FROM APPLICATION PROFILE\n")
    print("%s".rjust(len("%s") + 4) % "--> Getting list of EPGs in fabric. Hang on!")
    epg_df = pd.DataFrame.from_records(get_epgs(mo_dir))
    print("%s".rjust(len("%s") + 4) % "--> Getting list of Static Path Bindings. Almost there...")
    bind_df = pd.DataFrame.from_records(get_bindings(mo_dir))
    print()
    tn = choose_tn()
    app_p = choose_app_p()
    epg_df = epg_df[epg_df['appProfDn'] == ('uni/tn-%(tenant)s/ap-%(appProf)s' % dict(tenant=tn, appProf=app_p))]
    bind_df = bind_df[bind_df['appProfDn'] == ('uni/tn-%(tenant)s/ap-%(appProf)s' % dict(tenant=tn, appProf=app_p))]
    clear()
    return epg_df, bind_df


def get_epgs(mo_dir):
    fv_aepg = mo_dir.lookupByClass("fvAEPg")
    epg_list = [{"appProfDn": fv_aepg[idx]._BaseMo__parentDnStr, "epgDn": fv_aepg[idx].dn}
                for idx in range(len(fv_aepg))
                ]
    return epg_list


def get_bindings(mo_dir):
    rs_path_att = mo_dir.lookupByClass("fvRsPathAtt")
    bind_list = [{"appProfDn": re.sub(r'(uni/tn-.+/)(ap-.+)/(epg-.+)',
                                      r'\1\2', rs_path_att[idx]._BaseMo__parentDnStr),
                  "epgDn": rs_path_att[idx]._BaseMo__parentDnStr,
                  "bindTDn": rs_path_att[idx].tDn,
                  "encap": rs_path_att[idx].encap,
                  "instrImedcy": rs_path_att[idx].instrImedcy
                  }
                 for idx in range(len(rs_path_att))
                 ]
    return bind_list


def goodbye():
    print(colored(" GOOD BYE!\n", color="green", attrs=["bold", "blink"]))
    sleep(2)
    sys.exit()


def final_goodbye():
    print("\n" + "%s".center(70) % colored(" Thanks for using this interactive tool.",
                                               color="green", attrs=["blink"]))
    print("\n" + "%s".center(30) % colored(" Please provide your feedback to request new features or "
                                               "to report any problems at: ", color="green", attrs=["blink"]))
    print("%s\n".center(93) % colored("miguecor@cisco.com", color="blue", attrs=["blink"]))
    sleep(4)


def choose_tn():
    retries = 3
    for i in range(retries):
        try:
            tn_name = input("%s".rjust(len("%s") + 4) % colored("Enter the Tenant's name without the heading 'tn-': ",
                                                                color="magenta"))
            if tn_name == "":
                raise ValueError("Empty")
            elif tn_name.startswith("tn-"):
                raise ValueError("%s" % tn_name)
            else:
                return tn_name

        except ValueError as error:
            print(colored("%s" % f"\n '{error}' is not a valid Tenant name.\n"
                                 " Please provide a valid Tenant name without the leading 'tn-'", color="red"))
            print(colored(" You can try %s more time(s).\n", "yellow") % ((retries - 1) - i))
            continue
    goodbye()


def choose_app_p():
    retries = 3
    for i in range(retries):
        try:
            app_p_name = input("%s".rjust(len("%s") + 4) % colored("Enter the Application Profile's name: ",
                                                                   color="magenta"))
            if app_p_name == "":
                raise ValueError("Empty")
            elif app_p_name.startswith("ap-"):
                raise ValueError("%s" % app_p_name)
            else:
                return app_p_name

        except ValueError as error:
            print(colored("%s" % f"\n '{error}' is not a valid Tenant name.\n"
                                 " Please provide a valid Tenant name without the leading 'tn-'", color="red"))
            print(colored(" You can try %s more time(s).\n", "yellow") % ((retries - 1) - i))
            continue
    goodbye()


def create_bindings_from_excel():
    global count, auth_info
    try:
        if path.exists(excel_file):
            apic_id = re.sub(r"http(s)?://(.*)", r"\2", auth_info[0])
            print("\n" + "%s".rjust(len("%s") + 4) % colored(f" The '{excel_file}' file was found in the local "
                                                             f"directory."))
            print("%s".rjust(len("%s") + 4) % colored(" The script will proceed to deploy the Static Path Bindings "
                                                      "from the file to the following ACI fabric:"))
            print("\n" + "%s".rjust(len("%s") + 4) % colored(f" ACI Fabric IP/Hostname/FQDN: {apic_id}",
                                                             color="green", attrs=["bold"]))
            input("\n" + "%s" % colored(" Please hit 'Enter' when you are ready to start... ",
                                        color="magenta"))
            clear()
            epg_without_binding = pd.read_excel(excel_file)
            epg_without_binding['encap'] = ["vlan-%s" % value for idx, value
                                            in epg_without_binding['encap'].iteritems()]
            epg_without_binding = epg_without_binding.set_index('epgDn')
            cfg_request = ConfigRequest()
            for epg_dn, value in epg_without_binding[['tDn', 'encap']].iterrows():
                this_tn, this_app_p, this_epg = (re.sub(epg_regex, r'\2 \3 \4', epg_dn)).split()
                fv_rs_path_att_mo = RsPathAtt(epg_dn, tDn=value[0], instrImedcy="immediate", encap=value[1])
                cfg_request.addMo(fv_rs_path_att_mo)
                response = mo_dir.commit(cfg_request)
                print("\n" + "%s" % colored(f" Deploying on {this_epg}... ", color="blue", attrs=["bold"]))
                count += 1
                with open(filename, "a+") as file:
                    file.write("%s: %s fvRsPathAtt_dn: %s %s\n"
                               % (datetime.now().strftime("%Y/%m/%d %H:%M:%S"),
                                  apic_id, fv_rs_path_att_mo.dn, response))
                sleep(2)
            remove(excel_file)
            sleep(2)
        else:
            pass

    except Exception as error:
        print(colored("%s" % f"\n Error in function 'create_bindings_from_excel': {error}", color="red"))


def generate_excel(epg_without_binding, bind_tdn):
    try:
        if epg_without_binding.size == 0:
            print(colored("\n All EPGs have a Static Path Binding Associated :)\n", color="green", attrs=["bold"]))
            sleep(4)
        else:
            if path.exists(excel_file) is False:
                epg_without_binding['tDn'] = "%s" % bind_tdn
                epg_without_binding['encap'] = ""
                epg_without_binding.to_excel(excel_file, index=False)
                print("\n" + "%s" % colored(f" An '{excel_file}' file has been generated with the information of\n"
                                     f" the EPGs that don't have a Static Path Binding already configured.",
                                     color="yellow", attrs=["bold"]))
                print("%s" % colored(" Please open the file to edit the 'encap' column with the necessary VLAN "
                                     "encapsulation value.", color="yellow", attrs=["bold"]))
                print("\n%s %s\n" % (colored(" IMPORTANT: ", color="magenta", attrs=["bold"]),
                                 colored("Please make sure to enter only numeric values (e.g. 2, 50, 101, 2050)",
                                         color="yellow")))
                sleep(4)
                print(colored("%s", "blue") % " How would you like to proceed?\n\n",
                      "%s".rjust(len("%s") + 4) % "1) Wait until the Excel file has been filled and continue\n",
                      "%s".rjust(len("%s") + 4) % "2) Exit the script to fill the Excel file and try again later\n")
                selection = select(2)
                if selection == 1:
                    input("\n" + "%s" % colored(" Please fill the Excel file with the necessary information and \n"
                                         " hit 'Enter' when you are ready to proceed...", color="magenta"))
                    sleep(2)
                else:
                    print("%s" % colored(" Thank you.  Please come back when you are ready to continue.",
                                         color="yellow"))
                    sleep(2)
                    goodbye()

    except Exception as error:
        print(colored("%s" % f"\n Error in function 'generate_excel': {error}", color="red"))


def create_bindings_from_ready_df(binding_ready_to_deploy, bind_tdn):
    global count
    apic_id = re.sub(r"http(s)?://(.*)", r"\2", auth_info[0])
    try:
        binding_ready_to_deploy['size'] = ''
        binding_ready_to_deploy['size'] = [value.size for epg_dn, value in binding_ready_to_deploy.encap.iteritems()]
        binding_ready_to_deploy = binding_ready_to_deploy.sort_values(by='size', ascending=False)
        cfg_request = ConfigRequest()
        for epg_dn, value in binding_ready_to_deploy['encap'].iteritems():
            this_tn, this_app_p, this_epg = (re.sub(epg_regex, r'\2 \3 \4', epg_dn)).split()
            if value.size >= 2:
                clear()
                print(
                    "%s" % colored(
                        "\n The following EPG has Static Path Bindings with multiple encapsulation values:\n\n",
                        color="yellow") +
                    "          Tenant:    %(this_tn)s\n"
                    "     App Profile:    %(this_appProf)s\n"
                    "             EPG:    %(this_epg)s\n" % dict(this_tn=this_tn, this_appProf=this_app_p,
                                                                 this_epg=this_epg) +
                    "%s" % colored("\n Please select the encapsulation value to be deployed from the list below, "
                                   "or enter a new one:\n", color="yellow")
                )
                opt_dict = {}
                opt_idx = 1
                while opt_idx <= value.size:
                    for item in value:
                        opt_dict["%s" % opt_idx] = item
                        opt_idx += 1
                opt_dict["%s" % opt_idx] = "Enter a new encapsulation value"
                opt_idx = 1
                for opt in opt_dict.values():
                    print("%s".rjust(len("%s") + 4) % f" {opt_idx}) {opt}")
                    opt_idx += 1
                print()
                selection = select(value.size + 1)
                if opt_dict[str(selection)] == "Enter a new encapsulation value":
                    print("\n" + "%s" % colored(f" Please enter an encapsulation value for {this_epg}",
                                                color="magenta") +
                          "\n%s" % colored(" IMPORTANT: ", color="magenta", attrs=["bold"]) +
                          "%s\n" % colored("Please make sure to enter only numeric values "
                                           "(e.g. 2, 50, 101, 2050)", color="yellow")
                          )
                    encap_value = input("%s" % colored(" --> ", color="magenta"))
                    is_good_number = encap_value.isnumeric() and 1 <= int(encap_value) <= 4096
                    while is_good_number is False:
                        print("%s" % colored(" The provided encapsulation value '", color="yellow") +
                              "%s" % colored(f"{encap_value}", color="magenta", attrs=["bold"]) +
                              "%s" % colored("' is not numeric OR is zero OR is higher than 4096.", color="yellow"))
                        encap_value = input("\n" + "%s" %
                                            colored(f" Please enter an encapsulation value for {this_epg} --> ",
                                                    color="magenta"))
                        is_good_number = encap_value.isnumeric() and 1 <= int(encap_value) <= 4096
                    vlan_value = "vlan-%s" % encap_value
                    fv_rs_path_att_mo = RsPathAtt(epg_dn, tDn=bind_tdn, instrImedcy="immediate", encap=vlan_value)
                    cfg_request.addMo(fv_rs_path_att_mo)
                    print("\n" + "%s" % colored(f" Deploying on {this_epg}... ", color="blue", attrs=["bold"]))
                    response = mo_dir.commit(cfg_request)
                    count += 1
                    with open(filename, "a+") as file:
                        file.write("%s: %s fvRsPathAtt_dn: %s %s\n" % (datetime.now().strftime("%Y/%m/%d %H:%M:%S"),
                                                                       apic_id, fv_rs_path_att_mo.dn, response))
                    sleep(2)

                else:
                    vlan_value = opt_dict[str(selection)]
                    fv_rs_path_att_mo = RsPathAtt(epg_dn, tDn=bind_tdn, instrImedcy="immediate", encap=vlan_value)
                    cfg_request.addMo(fv_rs_path_att_mo)
                    response = mo_dir.commit(cfg_request)
                    print("\n" + "%s" % colored(f" Deploying on {this_epg}... ", color="blue", attrs=["bold"]))
                    count += 1
                    with open(filename, "a+") as file:
                        file.write("%s: %s fvRsPathAtt_dn: %s %s\n" % (datetime.now().strftime("%Y/%m/%d %H:%M:%S"),
                                                                       apic_id, fv_rs_path_att_mo.dn, response))
                    sleep(2)

            elif value.size == 1:
                fv_rs_path_att_mo = RsPathAtt(epg_dn, tDn=bind_tdn, instrImedcy="immediate", encap=value[0])
                cfg_request.addMo(fv_rs_path_att_mo)
                response = mo_dir.commit(cfg_request)
                print("\n" + "%s" % colored(f" Deploying on {this_epg}... ", color="blue", attrs=["bold"]))
                count += 1
                with open(filename, "a+") as file:
                    file.write("%s: %s fvRsPathAtt_dn: %s %s\n" % (datetime.now().strftime("%Y/%m/%d %H:%M:%S"),
                                                                   apic_id, fv_rs_path_att_mo.dn, response))
                sleep(2)

    except Exception as error:
        print(colored("%s" % f"\n Error in function 'creating_bindings_from_ready_df': {error}", color="red"))


def main():
    try:
        banner()
        welcome()
        auth_info = get_apic_info()
        mo_dir = login(auth_info)
        bind_tdn = "%s" % choose_binding_type()
        epg_df, bind_df = where_to_deploy(mo_dir)  # Any deployment option should always return these 2 DataFrames
        epg_without_binding = epg_df[~epg_df['epgDn'].isin([value[1] for idx, value in bind_df.iterrows()])]
        binding_ready_to_deploy = pd.DataFrame(bind_df.groupby(['epgDn']).encap.unique())
        generate_excel(epg_without_binding, bind_tdn)
        create_bindings_from_excel()
        create_bindings_from_ready_df(binding_ready_to_deploy, bind_tdn)
        print("%s" % colored(f" A '{filename}' file has been generated with the information\n"
                             f" of the deployed objects.", color="green", attrs=["bold"]))
        print("\n" + "%s\n" % colored(f" A total of {count} Static Path Bindings have been configured.",
                                      color="green", attrs=["bold"]))
        mo_dir.logout()
        thank_you()
        final_goodbye()

    except KeyboardInterrupt:
        print("\n" + "%s" % colored(" User interrupted execution of script :( ", color="blue", attrs=["bold"]))
        goodbye()


if __name__ == "__main__":
    main()