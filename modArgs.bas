Attribute VB_Name = "modArgs"
Option Explicit

'===========================================================================
'
' modArgs
'
'---------------------------------------------------------------------------
'
' Name:           modArgs
' Author:         Daniel Keep
' Contact:        shortcircuitfky@hotmail.com
'                 When you contact me, please put [VB] in the subject line
'                 otherwise I might not read it!
' Version:        1.8
' Last Modified:  4 August 2002
' Description:    Good old GetArgs has been replaced with ProcessCMDLine :'(
'                 Let's have one minutes silence... ... Ok, back to work ;)
' Changes    1.8: Changed name to modArgs, and changed the names of most of
'                 the methods.  This is to make it conformant to my new
'                 [system]Method() naming scheme.  Just incase you haven't
'                 worked it out, the methods in this module all have a
'                 prefix of arg.  Also, the SlashPath() has been removed,
'                 and placed in modFile.
'                 Finally, added argGetSwitchArg(), which allows you
'                 to finally build apps that read command-line arguments
'                 /really/ fast: check for the switch using
'                 argSwitchPresent(), and then get the value using
'                 argGetSwitchArg().  Dead easy ;)
'            1.7: Added the AddArgument(),RemoveArgument(), RebuildCmdLine()
'                 and SlashPath() methods.
'           1.6b: Modified it to add the original GetArgs()... Guess I got
'                 sentimental ;)
'           1.6a: Added a useage section... ;)
'            1.6: Well, I just broke compatibility by changing the way argc
'                 is used.  Previously, a value of 0& meant there was one
'                 argument (index 0).  Now, argc is the ACTUAL number of
'                 arguments like it shoulda been from the start.  Just hope
'                 I remember doing this later on... :S
'           1.5a: For the 'a' update, I just did some optimizations and
'                 general cleaning up before I post it onto Artifact and
'                 PSC.
'
'---------------------------------------------------------------------------
'
' Useage:
' =======
'
'  Probably should have included this in the original release :$
'  Anyway, here's the standard procedure for using modArgs:
'  First up, call ProcessCMDLine().  There's no need to specify any
'  arguments: if you don't specify anything it automatically grabs the
'  command line from Command$.
'  Next, once the call is complete, you can access your arguments from the
'  argv() array.  Remember that if you didn't specify any arguments, argv(0)
'  will be your executable path (ie: C:\MYAPP\MYAPP.EXE).  Everything after
'  that is an argument.  Also, argc is the number of elements in the array,
'  so if you don't pass ProcessCMDLine() anything, it will always be at
'  least 1 (because the EXE path is included).  Here's an example:
'
'    C:\MYAPP\MYAPP /arg1 arg2 "arg3 arg4" arg5
'
'  Would generate:
'
'    argc: 5
'    argv(0): "C:\MYAPP\MYAPP.EXE"
'    argv(1): "/arg1"
'    argv(2): "arg2"
'    argv(3): "arg3 arg4"
'    argv(4): arg5
'
'  You will notice that argv(3) is "arg3 arg4".  This is because arguments
'  surrounded by quotes are treated as one argument.
'
' Update (1.6b):
'
'  I figured that since I've added support for GetArgs() back into the
'  module (now you know where the name comes from), I figured I'd better
'  write some usage instructions ;)
'  Essentially, GetArgs() is just like ProcessCMDLine(), except that the
'  first two arguments are required, and are a string array and long.  The
'  array is for storing the arguments, and the long the count.  This is
'  great for custom implementations (I added it back because I was working
'  on a simple command-processor, and didn't want the result to overwrite
'  the command-line arguments).  So, if you had an array args() and a long
'  called count, and you wanted to process 'MyCommand arg1 arg2 "arg3 arg4",
'  you would use:
'
'    modGetArgs.GetArgs args, count, "My Command arg1 arg2 ""arg3 arg4"""
'
'  Easy.
'
' Update (1.7):
'
'  The new methods are dead use to use.
'  AddArgument() simply appends a new argument to the array to save you from
'  all the messy ReDim stuff.  Just pass it the contents of the argument,
'  and it's there.
'  RemoveArgument() just takes the index of the argument you want to remove,
'  and gets rid of it.
'  RebuildCmdLine() takes the existing array and builds a new command line.
'  It even quotes the strings with spaces to make sure they'll work
'  properly.  The one problem is that if you have an argument like this:
'
'    this argument has some "spaces" in it
'
'  Then it's going to end up being interpreted strangely.  But can't help
'  it I suppose... :P
'  SlashPath() just ensures that there's a backslash at the end of a path.
'
' Update (1.7a):
'
'  Ok, the changes are as follows:
'
'    ProcessCMDLine() --> argProcessCMDLine()
'    GetArgs()        --> argGetArgs()
'    IsSwitch()       --> argSwitchPresent()
'             **NEW**     argGetSwitchArg()
'    AddArgument()    --> argAdd()
'    RemoveArgument() --> argRemove()
'    RebuildCMDLine() --> argRebuildCMDLine()
'    SlashPath()          Removed
'
' As for the new method, argGetSwitchArg(), you use the exact same
' syntax as argSwitchPresent(), except that instead of returning true/false,
' it returns the argument immediately AFTER the specified switch (if it
' finds it.)  So, if you had the command-line:
'
'  C:\MYAPP.EXE somefile.ext --readmode readonly --readmode readwrite
'
' And you called argGetSwitchArg("--readmode"), it would return "readonly".
' Note that using this method, there is no way to get the second --readmode
' argument "readwrite", which is some cases is what you would want.  Perhaps
' in a future version, I'll change it so that it reads the list backwards ;)
'
'===========================================================================

'===========================================================================
' Yes, this is really bad coding (global variables), but it feels more...
' C++'sy this way ;)  If you're really paranoid, you could always move these
' to ByRef arguments of ProcessCMDLine, although that defeats the purpose of
' having your arguments easilly accessible...
Global argc&, argv() As String

'===========================================================================
' argProcessCMDLine()
'   A more streamlined version of the old GetArgs(), this method takes the
'   current command-line arguments (in Command$), and parses them into an
'   array (argv()$).  Please note that the first argument is ALWAYS the path
'   to the program's executable ala C++.  You can also override the default
'   arguments (and the exe path) by supplying your own arguments in [Args]
Public Function argProcessCMDLine(Optional ByVal Args As String)
Dim i&

' This is now just a wrapper for GetArgs.  Call GetArgs to do the
' processing.
argGetArgs argv, argc, Args

' All done!  Now you can access the command-line arguments from
' argv() and argc

End Function

'===========================================================================
' argGetArgs()
'   Here's ProcessCMDLine() with GetArgs() parameters...  Does the same
'   thing as ProcessCMDLine(), but you need to supply an array and long...
Public Function argGetArgs(ByRef argv() As String, ByRef argc As Long, _
 Optional ByVal Args As String)
Dim i&

' This is the temporary variable (duh).  We keep the 'processed'
' (ie mutilated) version of the command line in here.
Dim strArgTemp$

' This is used to store character positions gleaned from InStr() calls
Dim lngCharPos&

' Do we need to pull the arguments from Command$?
If Args <> "" Then
  strArgTemp = Trim$(Args)
Else
  strArgTemp = Trim$(Command$)
End If

' Do we want to set the first argument to the EXE path?
If Args = "" Then
  
  ' Resize the array
  ReDim argv(0&)
  argc = 1&
  
  ' Save the value
  argv(0&) = App.Path
  If Right$(argv(0&), 1&) <> "\" Then argv(0&) = argv(0&) & "\"
  argv(0&) = argv(0&) & App.EXEName & ".exe"
  
Else
  
  ' Nope.  Set argc to 0 so it works good-like
  argc = 0&
  
End If

'---------------------------------------------------------------------------
' This is here for my debugger.  You can remove these conditional
' statements quite safely...
#If DEBUG_ENABLED Then
  dbgPrintLine
  dbgPrintLine "Command-line arguments:"
  dbgPrint "  [0]: "
  dbgPrintLine argv(0&)
#End If

'---------------------------------------------------------------------------
' Right, here's the main loop.  What we do, is every time we find an
' argument, we strip it from strArgTemp.  Ergo, when all arguments have been
' processed, the string is empty.  Simple, huh? :P
Do Until strArgTemp = ""
  
  ' First, we check to see if we're dealing with a quoted argument
  ' (ie: "this has three spaces!")
  If Left$(strArgTemp, 1&) = Chr$(34) Then
    
    ' Yup; increase the array by one
    argc = argc + 1&
    ReDim Preserve argv(argc - 1&)
    
    ' Find the ending quote
    lngCharPos = InStr(2&, strArgTemp, Chr$(34&))
    
    ' IS there an ending quote?  If not, use the rest of the string
    ' (The +2 is there to negate the -2 below which is designed to
    ' avoid catching that last quote... which we aren't worried
    ' about... ;)
    If lngCharPos = 0& Then lngCharPos = Len(strArgTemp) + 2&
    
    ' Strip out the argument
    argv(argc - 1&) = Mid$(strArgTemp, 2&, lngCharPos - 2&)
    
    ' Now remove that argument from the temp var
    strArgTemp = LTrim$(Mid$(strArgTemp, lngCharPos + 1&))
    
    ' Log that argument
    #If DEBUG_ENABLED Then
      dbgPrint "[" & argc & "]: "
      dbgPrintLine argv(argc - 1&)
    #End If
    
  Else
    
    ' No quotes; expand array
    argc = argc + 1&
    ReDim Preserve argv(argc - 1&)
    
    ' Now, are there actually any more spaces?
    If InStr(1, strArgTemp, " ") <> 0& Then
      
      ' Yes.  But first, check to see if there's a quote in this
      ' argument
      If InStr(1, strArgTemp, Chr$(34)) <> 0 And _
       InStr(1, strArgTemp, Chr$(34)) < InStr(1, strArgTemp, " ") Then
        
        ' Yes.  First, extract up to the first quote
        lngCharPos = InStr(1&, strArgTemp, Chr$(34))
        argv(argc - 1&) = Left$(strArgTemp, lngCharPos - 1&) & Chr$(34)
        strArgTemp = Mid$(strArgTemp, lngCharPos + 1&)
        
        ' Next, find the closing quote
        lngCharPos = InStr(1&, strArgTemp, Chr$(34))
        
        ' Does it exist?
        If lngCharPos <> 0& Then
          
          ' Yes, extract up till that point
          argv(argc - 1&) = argv(argc - 1&) & Left$(strArgTemp, lngCharPos - 1&) & Chr$(34)
          strArgTemp = Mid$(strArgTemp, lngCharPos + 1&)
          
        Else
          
          ' No... just extract the rest of the string
          argv(argc - 1&) = strArgTemp
          strArgTemp = ""
          
        End If
        
      Else
        
        ' Nope.  Just find and extract up till the next space
        lngCharPos = InStr(1&, strArgTemp, " ")
        
        ' Now strip out the argument, and remove it from strArgTemp
        argv(argc - 1&) = Left$(strArgTemp, lngCharPos - 1&)
        strArgTemp = Mid$(strArgTemp, lngCharPos + 1&)
        
      End If
      
    Else
      
      ' Nope.  The rest of the string IS the last argument
      argv(argc - 1&) = strArgTemp
      strArgTemp = ""
      
    End If
    
    ' Log the argument
    #If DEBUG_ENABLED Then
      dbgPrint "  [" & argc & "]: "
      dbgPrintLine argv(argc - 1&)
    #End If
    
  End If
  
  ' Trim the command line
  strArgTemp = Trim$(strArgTemp)
  
Loop

' All done!  Now you can access the command-line arguments from
' argv() and argc
#If DEBUG_ENABLED Then
  dbgPrintLine
#End If

End Function

'===========================================================================
' argSwitchPresent()
'   This is a little something I wrote because I am (like most programmers)
'   oh so very lazy.  Just feed it a switch (like /l, /W), and it will look
'   for it in your argv() array.  If it finds it, it returns True (guess
'   what it returns if it DOESN'T find it.... FALSE!  Bet you didn't expect
'   that :p).  What's more, it also returns the array index of that switch
'   (for extra processing) in Position.
'   If that wasn't enough, it also supports pattern matching, so you can
'   search for special switches like "/i:*".

Public Function argSwitchPresent(ByRef Switch As String, _
    Optional ByRef Position As Long = 0, _
    Optional ByVal UseWildcard As Boolean = False) As Boolean
Dim i&

' Do we want to use pattern matching?
If UseWildcard = True Then
  ' Yup; start searching
  For i = 0& To argc - 1&
    ' Compare using the Like operator
    If argv(i) Like Switch Then
      ' Return true, and the position
      argSwitchPresent = True
      Position = i
      Exit Function
    End If
  Next
Else
  ' Nup; start searching
  For i = 0& To argc - 1&
    ' Compare using the = operator (ohlike, wow...)
    If argv(i) = Switch Then
      ' Return true, and the position
      argSwitchPresent = True
      Position = i
      Exit Function
    End If
  Next
End If

' If it got here, it ain't there, so return false
argSwitchPresent = False

End Function

'===========================================================================
' argGetSwitchArg()
'  Returns the argument immediately after the specified switch.  Switch
'  finding is done in the same way that argSwitchPresent() does it.
Public Function argGetSwitchArg( _
  ByRef Switch As String, _
  Optional ByRef Position As Long = 0, _
  Optional ByVal UseWildcard As Boolean = False _
) As String
Dim i&

' Do we want to use pattern matching?
If UseWildcard = True Then
  ' Yup; start searching
  For i = 0& To argc - 1&
    ' Compare using the Like operator
    If argv(i) Like Switch Then
      ' Is there a next argument?
      If (i + 1&) < argc Then
        ' Yup, return it
        argGetSwitchArg = argv(i + 1&)
        Position = i + 1&
      Else
        ' Nope... return -1 and ""
        argGetSwitchArg = ""
        Position = -1&
      End If
      Exit Function
    End If
  Next
Else
  ' Nup; start searching
  For i = 0& To argc - 1&
    ' Compare using the = operator (ohlike, wow...)
    If argv(i) = Switch Then
      ' Is there a next argument?
      If (i + 1&) < argc Then
        ' Yup, return it
        argGetSwitchArg = argv(i + 1&)
        Position = i + 1&
      Else
        ' Nope... return -1 and ""
        argGetSwitchArg = ""
        Position = -1&
      End If
      Exit Function
    End If
  Next
End If

' If it got here, it ain't there, so return nothing
argGetSwitchArg = ""
Position = -1&

End Function

'===========================================================================
' argAdd()
'  This little puppy adds a new argument to the array.  Easy as.
Public Function argAdd(ByVal Argument As String)

' First, redimension the array
ReDim Preserve argv(argc)
argc = argc + 1&

' Now, append the argument
argv(argc - 1&) = Argument

' Done

End Function

'===========================================================================
' argRemove()
'  This method will remove a specified argument, and collapse the array.
Public Function argRemove(ByVal Index As Long)

' First up, do we need to redim the array, or erase it?
If argc = 1 Then
  
  Erase argv
  argc = 0&
  Exit Function
  
Else
  
  ' Loop through the elements, putting them back one index
  Dim i&
  For i = Index + 1& To argc - 1&
    argv(i - 1&) = argv(i)
  Next i
  
  ' Now, redim the array
  argc = argc - 1&
  ReDim Preserve argv(argc - 1&)
  
  ' Done
  Exit Function
  
End If

End Function

'===========================================================================
' argRebuildCmdLine()
'  Rebuilds the command line from the current array.
Public Function argRebuildCmdLine() As String

' Ok, here we are going to loop through argv[], appending the arguments to
' the string.
Dim m_strBuffer$, i&

If argc > 0& Then
  
  If InStr(argv(i), " ") > 0& Then
    m_strBuffer = Chr$(34) & argv(i) & Chr$(34)
  Else
    m_strBuffer = argv(i)
  End If
  
End If

For i = 1& To argc - 1&
  
  If InStr(argv(i), " ") <> 0& And InStr(argv(i), Chr$(34)) = 0& Then
    m_strBuffer = m_strBuffer & " " & Chr$(34) & argv(i) & Chr$(34)
  Else
    m_strBuffer = m_strBuffer & " " & argv(i)
  End If
  
Next i

' Return the command line
argRebuildCmdLine = m_strBuffer

End Function
