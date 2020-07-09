# interactive_epg_static_binding_tool
## A script that provides an interactive menu to deploy Static Ports on EPGs in an ACI fabric.

When running the script, a series of options are available for selection.  Depending on the selected options, the following types of Static Ports can be deployed:

- Access Port
- Regular Port-Channel
- Virtual Port-channel

The script will provide a number of options on how to choose the EPGs to which the Static Ports will be deployed:
1) Deploy on all EPGs from an AppProfile
2) Deploy on all EPGs from a VRF
3) Deploy on same EPGs as another port binding
4) Deploy from a CSV file
5) Deploy on a single EPG

At this moment, only the first option is available.

There are 3 cases that have been covered in the script:
1) When an EPG has Static Ports all with the same encapsulation value.
    - In this case the script will take the encapsulation value from the EPG and deploy the new Static Port with the same encapsulation value.
2) When an EPG has Static Ports with at least 2 different encapsulation values.
    - The script will take all the different encapsulation values in the EPG and provide the user with a list of possible selections.
    - One of the options is to enter a new encapsulation value different from the ones in the list.
        - __IMPORTANT__: Use only numeric values.
3) When an EPG does NOT have any Static Ports configured.
    - In this case the script will generate an `EPGS_WITHOUT_BINDING.xlsx` Excel file.
        - The content of the file will have a column named 'encap'.
        - This column should be filled with the desired encapsulation value for the specific EPG.
        - __IMPORTANT__: Use only numeric values.
    - The script then will give the option to:
        1) Wait until the Excel file has been edited and saved.
            - After the file has been edited, hit enter to continue the execution of the script.
            - The script will load the information from the Excel file and use it to deploy the Static Ports with the desired encapsulation.
        2) Stop the execution of the script and re-run it after making the necessary changes to the Excel file.
            - After the file has been edited, run the script again from the terminal.
            - The script will ask for the APIC authentication information again and will try to locate the Excel file.
            - The script will load the information from the Excel file and use it to deploy the Static Ports with the desired encapsulation.
    - After successfully deployment of the EPGs, the Excel file will be deleted.
    

After the previous steps have been completed, a log file will be generated with the following name: 
`Static_Ports_Deployment_yyyymmdd_hhmmss.log`
The file will contain information of the day, the time, the ACI fabric, the Static Port and the result of each Static Port deployment.


### USAGE
Simply run the script using Python3:
`$ python interactive_epg_static_binding_tool.py`


### HELP
There's no help menu in this script.


### REQUIREMENTS
To install the necessary modules for the script to run: 
`$ pip install -r requirements.txt`


### IMPORTANT
If you receive an error message about not being able to run the 'pandas' module, you might need to install 'xz' using your OS package manager before creating your Python environment.

For macOS users, you can check this article:  https://binx.io/blog/2019/04/12/installing-pyenv-on-macos/