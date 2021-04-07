# octopus-report for tenanted environment

*To get report working, please install following nuget package (run it in Powershell in Administrator mode)*

**Install-Package Octopus.Client -source https://www.nuget.org/api/v2 -SkipDependencies**

Handling a lot of projects in combination with many tenants might be confusing and configuration might not be so easy. This report will help you to get data from Octopus API using Octopus.Client nuget library. The output is saved in report.xlsx file, on which you can apply filters.

Ignored projects and Ignored packages needs to be changed manually in the *report.ps1*

If you are missing any report, don't hesitate to create an issue or PR!
