# Yammer Archiver

This is a prototype for archiving Yammer/Viva Engage messages in a group.

It uses the Microsoft Power Automate connector for Viva Engage and its [Get messages in a group (V3)](https://learn.microsoft.com/en-us/connectors/yammer/#get-messages-in-a-group-(v3)) action to retrieve messages from a specified group. The resulting data can be downloaded as JSON files and viewed with the `Dynamics 365 and Power Platform Preview Programs.html` file.

The prototype is based on the article [How to get all messages in a Yammer group using Microsoft Flow](https://alextofan.com/2019/03/18/how-to-get-all-messages-in-a-yammer-group-using-microsoft-flow/).

## Usage

1. Use the [Import Package (Legacy)](https://learn.microsoft.com/en-us/power-automate/export-import-flow-non-solution#import-a-flow) feature in Power Automate to import the zip file in the Power Automate Flow folder.
2. Get the feed ID of the group you want to archive. You can find this in the URL of the group page in Yammer/Viva Engage.
3. Run the flow and provide the feed ID as group id when prompted.
4. After the flow has finished, open the run and expand the "Set variable varAllMessages as output" step.
5. In the "Outputs" section, click the "Click to download" link to download the JSON file containing the messages.
6. Save the file in a new folder with the group name as the folder name next to the `Example group` folder. The file name should be the group name followed by `Messages`. See the sample file in the `Example group` folder.
7. Do steps 4-6 for the "Set variable varAllReferences as output" step and add the word `References` to the file name.
8. Add the group name in the `groups-config.json` file.
9. Open the `Dynamics 365 and Power Platform Preview Programs.html` file with a http server (e.g. using the Visual Studio Code [Live Server](https://marketplace.visualstudio.com/items?itemName=ritwickdey.LiveServer) extension).

![Viewing the Example Group messages](ExampleGroupMessagesView.png)