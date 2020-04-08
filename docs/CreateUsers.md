# Create Users, Security Groups, and Assign Users to Security Groups 

    NOTE

    This script is provided as a stop-gap means to onboard users into Azure AD.  It assumes the accounts will be cloud-only identities amenable to managing Firstline employees.
    At enterprise scale, this process is managed through your Identity and Access Management team.  It is not required to use this script in order to leverage the other scripts
    contained in this repository, but it is assumed you will need to update the scripts to accomodate how your Firstline employees are onboarded.

Creates users and security groups defined in their CSV files.  Assigns the user to their security group.

## Steps

1. Update the CSV file [Users](../data/users.csv) with the users that will be created
2. Update the CSV file [Security Groups](../data/securityGroups.csv) with the security groups that will be created
3. From PowerShell, run the script [Create Users](../scripts/CreateUsers.ps1)

## Contributing

This project welcomes contributions and suggestions. Most contributions require you to agree to a Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us the rights to use your contribution. For details, visit https://cla.microsoft.com.

When you submit a pull request, a CLA-bot will automatically determine whether you need to provide a CLA and decorate the PR appropriately (e.g., label, comment). Simply follow the instructions provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact opencode@microsoft.com with any additional questions or comments.
