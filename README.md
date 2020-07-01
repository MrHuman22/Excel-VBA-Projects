# Excel-VBA-Projects
Contains useful custom functions and subs I've written for a calendar worksheet used to organise the University of Newcastle's Science and Engineering Challenge Team

## [AddEventAssistant](https://github.com/MrHuman22/Excel-VBA-Projects/blob/master/AddEventAssistant.vba)
* Allows the user to add more than one comma-separated value into a cell from a data validation list. Choosing the same option again removes it from the list
* Adapted from [Sumit Bansal](https://trumpexcel.com). The original sub allowed multiple comma-separated selections without repitition, but not removal
* Follows an experiment with dictionaries in Excel (which allow you to use dict.Exists(SomeKey) instead of requiring InStr)
