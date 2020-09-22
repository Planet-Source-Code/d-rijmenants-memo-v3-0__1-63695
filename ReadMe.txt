
MEMO README 

Description
-----------

Memo is a small program to manage, edit and print memos.
A memo file is a file, which can contain a large number of
memos. You can use different memo files and import or
export other memos.


Memo Files
----------

The first time you start Memo, the program will prompt you 
to crete a new memo file.

The last opened Memo file will also be the one that is opened
the next time you stat the program.

Added or changed memos are saved automatically on exit.

In the Memo menu, you can open another memo file, create a new
one or copy the current memo file for backup.

The Memos
---------

New memos are added at the end of the file. The last memo is shown
as the program starts. The date in the toolbar is that of the last
changes on that memo.

To protect a memo, you can lock it. Erasing or changing it is now
impossible.

You can export one single memo, stored as memo file, import such a
single memo, or import a complete memo file to your current memo
file. Select 'Open' in the 'Memo' menu for all these operations.

By marking a memo, or several memos, as important, this is detected
on opening, and these memos are shown automatically.


Navigation and Search
---------------------

To find a certain memo, just use F3 or click the Find icon in the
lower toolbar. You can use up to three keywords, seperated by a space.

With the [Home],[Page down],[Page up] or [End] leys you navigate
through the memos. If the Search function is active, these keys are
used to search through the memo file. You can enter a keyword in the 
toolbar and hit enter for each continued search. On large memo files,
you can abort the search by using the [ESC] key.

You can also click on the position indicator, on the bottom, to jump
arround in the memo file.


View
----

Select the 'view' menu to change font, fontsize and color combination.
The selected font and fontsize are also used when printing the memo.


Sharing memo files
------------------

You can store the memo files on a network server to share them with
other users who also have this program. The access is managed by
a lock file.

Each time a memo file is opened, a memo lock file (.mlc) is written in
the same directorie as the memo file (make sure everyone has write
permission on that folder). If the memo file is free, you can edit it
and save the changes. If a second user opens the same memo file,
he it will open as read-only. The program will notify him and ask him
weither he wants to be notified if the memo-file is free for editing.

To see which user has the current write access, you can enter your 
user name in the 'Extra' menu (this has to be done only once). Without
the username, only the ID is shown to the second user.

On exit, the lock status is updated. If a problem has occurred during
exit (network problems etc) the lock status is not updated and the
memo file stays marked as locked. In that case, you can override the
lock status. This should only be done when you are sure that no other
user is currently using that memo file.


Splash screen
-------------

You can show a splash screen on start. This is done by simply adding
a memo that starts with the text [SPLASH]. All following text on that
memo is shown in the splash screen. 


Shortcut Keys
-------------

   Navigation

   F3  Find on/off
   Esc Find off or abort search


   MEMO-FUNCTIES:
   
   Ctrl+O Open memo file screen
   Ctrl+N Add new memo
   Ctrl+P Print current memo
   Ctrl+L Lock/unlock memo
   Ctrl+D Delet current memo
   Ctrl+Q exit program
   Ctrl+F Search function on/off (Find)


   (c) D Rijmenants 2002
