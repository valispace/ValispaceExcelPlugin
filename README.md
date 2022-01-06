# Valispace Excel Plugin

To install the Valispace Excel Plugin download the ValispaceExcelAddon.xlam file from the latest [releases](https://github.com/valispace/ValispaceExcelPlugin/releases "ExcelPlugin Releases") and follow these instructions:

1) Store the file in a folder of your choice.  
2) Open Excel, click on the main menu (or `File` in newer Excel versions) then `Options` --> `Add-Ins` --> Excel-Addins: `Go...`  
3) Click `Browse` and select the .xlam file, select it with a check-box and click `OK`  
4) A new Ribbon called `Valispace` should have appeared in the top main menu.  
5) To prevent the need to re-do the same steps every time you start Excel, you will have to add the folder containing the Add-In to your "Trusted File Locations". To do this, click on `File` --> `Options` --> `Trust Center` --> `Trust Center Settings...` --> `Trusted Locations` --> `Add new location...` --> `Browse` and add the folder where you stored the Addon.  
6) To setup Valispace, in the new `ValiSpace` Ribbon, select `Settings` and insert your deployment's URL (e.g. `https://yourdeployment.valispace.com`), your username and password and confirm with `Save`

Remember to use http**s**:// in the URL if you want the connection to be securely encrypted.

## Contribute

The source files of the plugin are in the src folder. To contribute to the plugin, edit those files and create a new package file before submitting a pull request.

## Version

The current version of the Valispace Excel Plugin is 1.2.
