Attribute VB_Name = "modvbTips"
'Interessante routines

'Give the Inet control the time it needs to execute commands
'When you use the Inet control to access an FTP site, you'll often
'need to issue commands in pairs. For instance, to change from one
'directory to another and then display the new directory's structure,
'you 'll need to first issue the CD (change directory) command, followed
'by the DIR (directory) command.

'However, in Visual Basic, if you try to execute these commands one
'right after another, the Inet control will most likely return an error.
'That 's because, depending on the connection speed used to access the
'FTP site, Visual Basic issues the second command before the FTP site
'has had a chance to complete the first command.

'To avoid this error, you'll need to wait until the FTP site has finished
'processing one command before continuing on to another. To do so, you
'can test the Inet control's StillExecuting property. As you can guess,
'this property indicates whether or not the last command is still
'executing. It returns True if so, and False if not. To take full
'advantage of this property, set it up in a loop in conjunction with
'the DoEvents keyword, like so:

'Private Sub myButton_Click()
'With Inet1
'   .Execute , "CDUP"
'   Do While .StillExecuting
'      DoEvents
'   Loop
'   .Execute , "DIR"
'End With
'End

