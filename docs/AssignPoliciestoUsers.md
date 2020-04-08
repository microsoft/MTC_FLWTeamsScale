# Assign Policies to Users

Assign the Teams policies for messaging, App Setup, and App Permissions according to their security group CSV file.  

## Steps

1. In the Teams Admin Console, create App Setup and App Permission policy for both the Firstline Manager and Worker 
2. Update the CSV file [Security Groups](./data/securityGroups.csv) for the assigned policies
3. From PowerShell, run the script [Assign Policies to Users](./scripts/AssignPoliciestoUsers.ps1)

## Contributing

This project welcomes contributions and suggestions. Most contributions require you to agree to a Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us the rights to use your contribution. For details, visit https://cla.microsoft.com.

When you submit a pull request, a CLA-bot will automatically determine whether you need to provide a CLA and decorate the PR appropriately (e.g., label, comment). Simply follow the instructions provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact opencode@microsoft.com with any additional questions or comments.
