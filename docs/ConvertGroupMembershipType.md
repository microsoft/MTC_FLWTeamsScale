# Convert Group Membership Type

When licensed for Azure AD P1 or above, you have the option of using Dynamic Group Membership vs. using assigned membership.  The scripts that created the Teams also created Office Groups of the membership type Assigned which means its members must be explicitly added.  Using Dynamic membership, rules
are written to determine if someone is a member of the team or not.

    NOTE

    When you run this script, the current membership of the group will be removed (except for its owners), 
    and new members will be added when the membership synch job runs.

## Steps

1. Update the CSV file [Migrate Groups](../data/migrateGroups.csv) with the Groups that will be migrated along with the rule for dynamic membership
2. From PowerShell, run the script [ConvertGroupMembershipType.ps1](../scripts/ConvertGroupMembershipType.ps1)

## Contributing

This project welcomes contributions and suggestions. Most contributions require you to agree to a Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us the rights to use your contribution. For details, visit https://cla.microsoft.com.

When you submit a pull request, a CLA-bot will automatically determine whether you need to provide a CLA and decorate the PR appropriately (e.g., label, comment). Simply follow the instructions provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact opencode@microsoft.com with any additional questions or comments.
