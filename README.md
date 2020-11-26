# D365 Power Kanban

This is a custom PCF control for displaying a Kanban inside D365 datasets.

Features so far:
- Supports option sets, status codes and boolean attributes as swimlane sources
- Color schemes are taken from option set color values
- Primary and Secondary entities can be configured, so that you can show accounts and their child cases in one kanban board (choose view Advanced)
- Swimlanes with Drag and Drop functionality (allowed transitions can be defined for statuscode attributes inside their respective state transitions, or via custom script hook)
- Record Count as badge per lane
- Card forms can be used for customizing the display, they are rendered with clickable lookups
- Side-By-Side view for quick data view AND edit, where all your form scripts are considered (resizable in lower right corner)
- Modal forms for data view and edit (open per double click on tile or "Open in modal" button)
- Custom Buttons and Custom Forms on state transition (via custom script support)
- Search in data, highlighting and filtering
- Support for subscriptions and notifications to records, so that you can subscribe to record changes and see them
- Custom Dialogs, for allowing to also resolve cases, win opportunities etc from the board
- Works on entity lists as well as on form subgrids
- Columns that contain html (such as email descriptions) are converted to plain text

Todos:
- Virtualization and Performance Tweaking

# Impressions
## Simple View (Primary Entity only)
![image](https://user-images.githubusercontent.com/4287938/90894070-55be9200-e3c0-11ea-822c-9674fd1cab89.png)


## Advanced View (Primary Entity and Secondary Entity)
![image](https://user-images.githubusercontent.com/4287938/90894139-71299d00-e3c0-11ea-9e0b-441187c1d224.png)


## Form View
![image](https://user-images.githubusercontent.com/4287938/90894207-856d9a00-e3c0-11ea-855c-48ea820e57ea.png)


## Custom Dialogs
![image](https://user-images.githubusercontent.com/4287938/90894247-96b6a680-e3c0-11ea-8f52-a5c12b402aa5.png)


## Side By Side View
![image](https://user-images.githubusercontent.com/4287938/90894676-468c1400-e3c1-11ea-9089-45ca0cca9f9c.png)


## Open in Modal
![image](https://user-images.githubusercontent.com/4287938/90894901-94088100-e3c1-11ea-9c00-cd9cb65d3420.png)


## Search with highlighting
![image](https://user-images.githubusercontent.com/4287938/90894995-b7cbc700-e3c1-11ea-8ed2-4373dec6ad7c.png)


## Notification List
![image](https://user-images.githubusercontent.com/4287938/90894792-6e7b7780-e3c1-11ea-8520-2629a9849732.png)

## Global Notifications
![image](https://user-images.githubusercontent.com/4287938/100028074-5e26ce80-2dee-11eb-9522-315ecaafa02c.png)

# Installation
- Download and install the [latest managed solution](/../../releases/latest)
- If you update from a version prior to 0.9.18 or want to use the same fields as in the Wiki examples, and you need the oss_taskid / oss_accountid / oss_incidentid fields on notification / subscription or the example script resource for an incident kanban, you can install the [D365PowerKanbanSamples](/../../tree/master/export) solution. They are not part of the default solution anymore. You can find the solution [here](/../../tree/master/export)

# Configuration
- Please see our [Wiki](https://github.com/XRM-OSS/D365-PowerKanban/wiki/Configuration)
