README -- EventLogger Example Program
-------------------------------------

Caveats
-------

This project was not the easiest in the world to get through, and I'm
putting this warning out right now: this is advanced stuff. However, the payoff
I believe is worth it: to produce a much more professional application
in Visual Basic, which has all of the beefed up Event logging capabilities
of its C++ cousin.

This project uses not only VB, but some C++ compiler tools as well to generate
special DLLs. The Zip file MsgDLL.zip contains a sample message file, and a
batch file, which will compile the DLL for you and put it in the appropriate
area.

The NTEventLog class is one of my own devising. There is very little that
one needs to do to work with it, other than have Windows NT or 2000 handy.
Please, look at the source for the example program before you run it. It could definitely use a little more work, such as parametric message handling, but that is for the future.

Included also is a Word document describing the process for creating a Message DLL file. It just describes *how* to make one. If I get enough hits back on this project, I'll put together something a little more meaty in terms of describing what you do after that. For the most part, you can look at the code in the NTEventLog class and understand the registry portion of the action.

Running The Example
-------------------

First, you must make sure you have Visual C++ on your machine. Then, unzip MsgDLL.zip, and run build.bat. This will create the appropriate DLL for the example program to use. 

Second, fire up VB and load the example project. If you are really curious, step through the code. There isn't much there, as there isn't much needed.

BTW, you can have the Event Viewer open at the same time as the program. It writes to the Application Log. If you do have it open simultaneously, you'll have to hit Refresh to see the messages come in.

Thanks
------

Special Thanks to Don Kiser for his excellent Registry Class, included in this project. Without it, this project simply wouldn't work.

And thanks also to Ian Ippolito, for Planet Source Code. Without it, I don't think I would have been inspired to do this.

