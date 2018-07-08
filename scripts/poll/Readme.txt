PacPoll 4.0 is simple to use. 

Version 4.0 now has 2 databases. Depending on what version of Access you use, you now have a choice. If you use Access 2000 then use the database poll.mdb. If you use an earlier version then rename poll97.mdb to poll.mdb.

PacPoll 4.0 does require the server to be using at least version 5 of the scripting engine. Nearly all host are now using at least 5.1, so this should not be a problem.

First at the top of the page you want to use the poll on you must place this line:
<%Response.Buffer=True%>

Second, drop the poll directory in the root of your web. On the page you want to have the poll, place this line where the poll should go:
<!-- #include virtual="poll/poll.asp" -->

Make sure the server has write permisions to the database and thats it for the installation.

Now to change the poll, open a browser to http://www.yoursite.com/poll/admin/login.asp

Log in as admin with password admin123456. As soon as you log in, you will find at the top of the page a form to change the Username and password. DO THIS FIRST!

For security reasons I recommend changing the name of the admin directory. This will keep inscrupulous people from playing with your settings. I learned this myself the hard way.

Then simply fill out the form. You can change the question and the four responses. You can have 2 3 or 4 responses if you only want two then just leave 3 and 4 blank. There are 30 fonts to choose from and you can set the color of the font as well as the size. You can also set the color of the bar graph and poll background as well.

One last thing. PacPoll 4.0 sets a cookie so that a visitor can only answer the poll once. When you change the poll it changes the cookie info as well so that when the new poll is set returning visitors will be able to answer the new poll.

Thats it you're done.


