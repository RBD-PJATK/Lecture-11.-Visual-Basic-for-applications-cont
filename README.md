# Lecture 11. Visual Basic for applications (cont.)

<h3>Abstract</h3>

<p>Lecture 11 further elaborates on programming database applications. We
show built-in objects of VBA which represent commands
(<i>DoCmd</i>), errors (<i>Err</i>), the application (<i>Application</i>)
and debug facilities (<i>Debug</i>).</p>

<p>We present some programming techniques like error handling, the requery
of a combo box, event cancelation, the execution of an SQL statement
and the code for multiple selection in a list box.</p>

<hr><h3><a name="Lista">Built-in objects</a></h3>

<p>Many commands of VBA are methods of the objects of a user interface as
well as the following built-in objects.</p>

<dl>
<dt><i>DoCmd</i>
<dd>It executes commands.
<dt><i>Application</i>
<dd>It represents the MS Access application.
<dt><i>Debug</i>
<dd>It represents the debug environment.
<dt><i>Err</i>
<dd>It represents errors.
</dl>

<hr><h3><a name="DoCmd">DoCmd</a></h3>

<h4>GoToRecord</h4>

<p>The <code>GoToRecord</code> method navigates to another record. For example the following
call initiates navigation to the first record.</p>

<pre>DoCmd.GoToRecord , , acFirst</pre>

<p>Here is the list of named constants that can be used as arguments
of <code>GoToRecord</code>.

<ul>
<li><code>acFirst</code>
<li><code>acPrevious</code>
<li><code>acNext</code>
<li><code>acLast</code>
<li><code>acNewRec</code> (navigates to a new empty record)
</ul>

<h4>OpenForm</h4>

<p>The <code>OpenForm</code> method opens the specified form. 
It has the following syntax (brackets mark optional arguments).</p>

<pre>DoCmd.OpenForm <i>FormName</i>, [<i>View</i>], [<i>FilterName</i>], _
               [<i>WhereCondition</i>], [<i>DataMode</i>], [<i>WindowMode</i>]</pre>

<p>Some of the optional arguments have default values.</p>

<dl>
<dt><code><i>View</i></code> = <code>acNomal</code>
<dd>A form will be opened in the form view.
<dt><code><i>DataMode</i></code> = <code>acFormPropertySettings</code>
<dd>The setting is taken from the property sheet of the form.
<dt><code><i>WindowMode</i></code> = <code>acWindowNormal</code>
<dd>An ordinary window will be opened.
</dl>

<p>In order to open the <i>Employees</i> form, you can write the following command.</p>

<pre>DoCmd.OpenForm "Employees", , , , acFormEdit</pre>

<p>In this command, second, third, fourth and sixth arguments are not specified.
The method will assume default values for them.  Empty spaces between commas mean
that the corresponding argument is not passed.  Note that we can totally omit the
sixth argument because there are no arguments after it.</p>

<p>Instead of <code>acFormEdit</code> you can use other named constants
<code>acFormAdd</code>, <code>acFormPropertySettings</code> and <code>acFormReadOnly</code>.
</p>

<h4>OpenReport</h4>

<p>The <code>OpenReport</code> method opens the specified report.  It is similar to 
<code>OpenForm</code> but has fewer arguments. 
It has the following syntax (brackets mark optional arguments).</p>

<pre>DoCmd.OpenReport <i>ReportName</i>, [<i>View</i>], [<i>FilterName</i>], [<i>WhereCondition</i>]</pre>

<h4>Close</h4>

<p>The <code>Close</code> method closes the specified object or the currently active
object if no object is indicated.  In order to close the active object, run the following
command.</p>

<pre>DoCmd.Close</pre>

<p>If you want to close the <i>Employees</i> form, use:

<pre>DoCmd.Close acForm, "Employees"</pre>

<h4>PrintOut</h4>

<p>The <code>PrintOut</code> method prints the currently active object.  It can be
a form, a report, a datasheet, a data access page and a module.</p>

<pre>DoCmd.PrintOut</pre>

<h4>OpenQuery</h4>

<p>The <code>OpenQuery</code> method opens the specified query.</p>

<pre>DoCmd.OpenQuery <i>QueryName</i>, acNormal, acEdit</pre>

<h4>RunSQL</h4>

<p>The <code>RunSQL</code> method executes the specified SQL statement.</p>

<pre>DoCmd.RunSQL <i>SQLStatement</i></pre>

<h4>Requery</h4>

<p>Method <code>Requery</code> repeats the query that is the row source for a form
or a list box.</p>

<p>For the list box <i>Deptno</i> on the current form, use:</p>

<pre>DoCmd.Requery "Deptno"</pre>

<p>However, if the list box does not belong to the current form, we have to
use the <code>Requery</code> method of this list box:</p>

<pre>Forms![Employees]![Deptno].Requery</pre>

<p>To requery the data of the current form, use:</p>

<pre>DoCmd.Requery</pre>

<p>If a form is not current, we have to use the <code>Requery</code>
method of this form:</p>

<pre>Forms![Employees].Requery</pre>

<h4>Quit</h4>

<p>The <code>Quit</code> method closes the application.</p>

<pre>DoCmd.Quit</pre>

<p>You can also use the same method of the <code>Application</code> object.</p>

<pre>Application.Quit</pre>


<h4>ApplyFilter</h4>

<p>The <code>ApplyFilter</code> method filters the contents of the current form. For example
we can filter books with respect to an unbound item <i>[From Date]</i>:</p>

<pre>DoCmd.ApplyFilter , "[Published] >= Forms![Search books]![From Date]"</pre>

<h4>ShowAllRecords</h4>

<p>The <code>ShowAllRecords</code> method displays all records on the current form.</p>

<pre>DoCmd.ShowAllRecords</pre>

<h4>Filtering records</h4>

<p>Instead of methods <code>ApplyFilter</code> and <code>ShowAllRecords</code>
of object <code>DoCmd</code> you can use the following properties of the form.</p>

<dl>
<dt><code>Filter</code>
<dd>The filtering condition.
<dt><code>FilterOn</code>
<dd>Logical values that indicated whether the filtering is on.
</dl>

<p>Inside the module of a form, you can switch filtering on with this command:</p>

<pre>Me.FilterOn = True</pre>

<p>If you want to switch it off (i.e. <code>ShowAllRecords</code>), use:</p>

<pre>Me.FilterOn = False</pre>

<p>If you set a filter in the form view and then close it, this filter will
be stored in its <code>Filter</code> property.   When you open this form again,
this filter will be available but switched off (<code>FilterOn = False</code>).
In order to activate, you have to change this property.</p>

<p>It is time for an exercise.</p>

<p align="center">
<table><tr><td class="notec">
Build a form based on a single table.  The form shall contain a set of buttons that
perform:

<ol>
<li>navigation among records, 
<li>navigation to a new empty record, 
<li>deletion of the current record,
<li>printing the contents of the table,
<li>closing the form,
<li>setting the filter by means of the <code>Me.Filter</code> property,
<li>switching the filter on and off by means of the
	 <code>Me.FilterOn</code> property,
<li>displaying the on-line help.
</ol>

</table> 

<h4>RunCommand</h4>

<p>The <code>RunCommand</code> method executes items of the built-in menus.</p>

<pre>DoCmd.RunCommand <i>Command</i></pre>

<p>You may omit object <code>DoCmd</code> and write simply:</p>

<pre>RunCommand <i>Command</i></pre>

<p><code>Command</code> is one of the named constants that correspond to the menu items.
The list of these constants can be found in the on-line 
help of MS Access under the subject
<i>RunCommand Method Constants</i>.</p>

<p>For example, the following command executes the <i>Options</i> item 
from the <i>Tools</i> menu.</p>

<pre>DoCmd.RunCommand acCmdOptions</pre>

<p>Here is another exercise for you.</p>

<table><tr><td class="notec">
Repeat the previous exercise, but program the operations on records by means of
the <code>RunCommand</code> method.</table> 

<hr><h3><a name="Debug">Debug</a></h3>

<p>The <code>Debug</code> object represents the debug facilities of MS Access.</p>

<h4>Print</h4>

<p>The <code>Print</code> method displays control information in the <i>Immediate Window</i>,
e.g.</p>

<pre>Debug.Print "I start calculating ROE"</pre>

<hr><h3><a name="Instrukcja">MsgBox</a></h3>

<p>The <code>Debug.Print</code> method is intended for developers of
applications as a tool to debug code.  On the other hand, there are also
commands which interact with the end user. One of them is
<code>MsgBox</code> which displays information for the user and a number of
simple buttons (OK, Cancel, Yes. No).  Their set depends on the arguments
passed to <code>MsgBox</code>.  This command has the following syntax.</p>

<pre>MsgBox <i>Prompt</i>, <i>Buttons</i>, <i>Title</i></pre>

<dl>
<dt><code><i>Prompt</code></i>
<dd>The message to the user.
<dt><code><i>Buttons</code></i>
<dd>The specification of buttons to appear as well as the icon, the default button and 
the kind of the window (modal or not).  The default for this argument
is <code>vbOKOnly</code> (equal to zero). 
<dt><code><i>Title</code></i>
<dd>The title of the message window.
</dl>

<p>For example, this command displays the message windows shown below.</p>

<pre>MsgBox "Invalid password.", , "ACCESS DENIED"</pre>

<p align="center"><img src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad11/images/11_1.png"></p>

<p>Such direct interaction with the end user is necessary if the application
notifies him/her that an error occurred.</p>

<p>The next command asks for confirmation by means of the following dialog.</p>

<pre>x = MsgBox ("Save before exit?", vbYesNo, "WARNING")</pre>

<p align="center"><img src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad11/images/11_2.png"></p>

<p>The actions that follow usually depend on the value returned by <code>MsgBox</code>,
e.g.</p>

<pre>x = MsgBox ("Save before exit?", vbYesNo, "WARNING")
If x = vbYes Then
  ' Save
  ...
End If</pre>

<p><code>MsgBox</code> may also be used when some data is lacking and the
application must ask the user to enter missing items.  In the following
example you find the event procedure for <i>Before Update</i> that will
display a message, when the mandatory field <i>Deptno</i> is empty.  In this
case, the procedure cancels the update. The user has only one choice, he/she
must accept the message by clicking the button OK.</p>

<pre>Private Sub Deptno_BeforeUpdate(Cancel As Integer)

  Dim strMsg As String, strTitle As String
  Dim intStyle As Integer
  strMsg   = "You have to assign this employee to a department."
  strTitle = "Missing department"
  intStyle = vbOKOnly + vbExclamation

  If IsNull([Deptno]) OR [Deptno] = "" Then
    MsgBox strMsg, intStyle, strTitle
    Cancel = True
  End If

End Sub</pre>

<hr><h3><a name="InputBox">InputBox</a></h3>

<p>VBA also provides the <code>InputBox</code> command which prompts the user for data.
It displays a dialog with the text box for the text to be entered by the user, e.g.</p>

<pre>Subject = InputBox("What is the subject of the book you search?")</pre>

<hr><h3><a name="Bezpo">Error handling</a></h3>

<p>Imagine a button that moves the user to the previous record.  Its event procedure
contains the following command.</p>

<pre>DoCmd.GoToRecord , , acPrevious</pre>

<p>When the current record is the first one and the user presses this button,
this command cannot be executed properly and an error will occur.
Such errors can be detected and handled. You declare the error handler by means of the
following statement.</p>

<pre>On Error GoTo <i>Error_handler</i></pre>

<p><i>Error_handler</i> is a label in the procedure.  In case of error control
is passed to the command that follows this label. Please take a look at an 
example of the event procedure for our button which contains an event handler.</p>

<pre>Private Sub Move_Previous_Click ()
On Error GoTo Error_handler

  DoCmd.GoToRecord , , acPrevious
  Exit Sub

Error_Handler:
  MsgBox "This is already the first employee!"
  Resume Next

End Sub</pre>

<p>The <code>Resume Next</code> command brings control back to the place
where the error occurred, skips the erroneous command and <i>resumes</i>
the execution at the <i>next</i> command.</p>

<h4><a name="Err">Err</a></h4>

<p>The built-in object <code>Err</code> represents the error that occurred during
the execution
of VBA code. It has several properties:</p>

<dl>
<dt><code>Err.Number</code>
<dd>The number of the most recent error.
<dt><code>Err.Description</code>
<dd>The description of the most recent error.
<dt><code>Err.Raise <i>Error_number</i></code>
<dd>The method that raises the error with number <i>Error_number</i>.
	This can be used e.g. in error handling, when we convert the caught error to
	another error.
</dl>

<p>We can also check the description of an error by means of the <code>Error</code>
function.
The following expression returns the description of error 657.</p>

<pre>Error(657)</pre>

<h4>On Error</h4>

<p>The <code>On Error</code> declaration tells the system what to do when an error occurs.
We have several options.</p>


<dl>
<dt><code>On Error GoTo <i>Error_handler</i></code>
<dd>Handle error by execution of the commands after a label <code><i>Error_handler</i></code>.
<dt><code>On Error Resume</code>
<dd>Repeat the command that caused the error.
<dt><code>On Error Resume Next </code>
<dd>Skip the command that caused the error and resume at the next command.
<dt><code>On Error Resume <i>Resume_label</i></code>
<dd>Resume the execution at label <code><i>Resume_label</i></code>.
<dt><code>On Error Stop</code>
<dd>Stop the execution.
<dt><font color="magenta"><code>On Error GoTo 0</code></font>
<dd>Cancel the previously declared method of error handling.
</dl>

<p><code>On Error GoTo <i>Error_handler</i></code> declares the switch to error handling
mode. Other options handle the error immediately and do not send control to error
handling code.  If the error handler is executing, the error handling mode can be exited 
in one of the following ways.</p>


<dl>
<dt><code>GoTo <i>Resume_label</i></code>
<dd>Go to label <i>Resume_label</i></code>.
<dt><code>Resume</code>
<dd>Repeat the command that caused the error.
<dt><code>Resume Next</code>
<dd>Skip the command that caused the error and resume at the next command.
<dt><code>Resume <i>Resume_label</i></code>
<dd>Resume the execution at label <code><i>Resume_label</i></code>.
<dt><code>Stop</code>
<dd>Stop the execution.
</dl>


<hr><h3><a name="Aktualizacja">Update of a combo box</a></h3>

<p>A combo box is a text box combined with a drop-down list. A user can either
select a value from this list or type a value into the text box.  We can limit
users' freedom by setting the combo box's property <i>Limit To List</i>.</p>

<p>If this property is set, the user types anything that does not belong to
the list and the combo box is exited, the <i>On Not in List</i> event will occur.
The default handler of this event notifies the user and disallows exiting the field.
</p>

<p>The event procedure for this event has two arguments.</p>

<dl>
<dt><code>NewData</code>
<dd>This is an input argument that passes the value entered by the user.
<dt><code>Response</code>
<dd>This is an output argument that tells the application what to do with the value
	entered by the user. There are three possibilities:

<dl>
<dt><code>acDataErrAdded</code>
<dd>Does not display a message to the user but enables you to add the entry to 
	the combo box list in the <i>NotInList</i> event procedure. 
<dt><code>acDataErrContinue</code>
<dd>Does not display the default message to the user. You can use this when you
	want to display a custom message to the user. For example, the event procedure
	could display a custom dialog box asking if the user wanted to save the new entry.
	If the response is <i>Yes</i>, the event procedure would add the new entry
	to the list and set the <code>Response</code> argument to
	<code>acDataErrAdded</code>. If the response is <i>No</i>,
	the event procedure would set the <code>Response</code> argument
	to <code>acDataErrContinue</code>.
<dt><code>acDataErrDisplay</code>
<dd>Use the default method, i.e. notify the user and reject the value.
</dl>

</dl>

<h4>Example 1</h4>

<p>We are building an <i>Employees</i> form which will be used to enter data on employees. 
It contains a combo box <i>Jobs</i>.  The list of jobs is hard-coded into the combo box
as the value of property <i>Row Source</i>. We have already set property <i>Limit To List</i>
for this combo box. Now we are going to code the event procedure.</p>

<pre>Private Sub Jobs_NotInList(NewData As String, Response As Integer)
   Dim ctl As Control
   Set ctl = Me!Jobs
   If MsgBox("Job not in list. Do you want me to add it?", vbOKCancel) = vbOK Then
      Response = acDataErrAdded
      ctl.RowSource = ctl.RowSource &amp; ";" & NewData
   Else
     Response = acDataErrContinue
     ctl.Undo
   End If
End Sub</pre>
 
<p>At the beginning we assign the reference to control <i>Jobs</i> to variable
<code>ctl</code>.  In VBA all controls are objects of the <code>Control</code>
class.</p>

<p>In the conditional statement we ask the user to confirm addition of a new job.
If the user confirms, we set <code>Response</code> to
<code>acDataErrAdded</code>, i.e. we notify MS Access that we added the value 
of argument <code>NewData</code> to the row source.</p>

<p>If the user chooses <i>Cancel</i>, we <code>Response</code> to
<code>acDataErrContinue</code>, i.e. we make MS Access continue without the default
error message. Next, we clear the value entered by the user (<code>clt.Undo</code>).</p>

<p><b>Warning!</b> This change of the <i>Row Source</i> property applies only 
to the current instance of the form and is not saved permanently in the property sheet
of the form.</p>

<p>Now please do the following exercise.</p>

<p align="center">
<table><tr><td class="notec">
Check the sequence of events <i>On Not in List</i>, <i>Before Update</i>,
<i>On Exit</i> in the case of a combo box.
</table> 

<hr><h3><a name="Procedury">Parameters of event procedures</a></h3>

<p>As you have probably noticed, some event procedures have parameters. If 
you create such a procedure, MS Access automatically generates the skeleton 
of the procedure with the names of these parameters.
We describe these procedures here, but we have skipped the procedures
for filter events (<i>On Filter</i> and <i>On Apply Filter</i>).</p>

<p>The two fundamental parameters of event procedures are 
<code>Cancel</code> and <code>Response</code>.</p>

<dl>
<dt><code>Cancel</code>
<dd>If you set it to <code>True</code>, the event is canceled.
<dt><code>Response</code>
<dd>Usually its value is one of the following:

<dl>
<dt><code>acDataErrContinue</code>
<dd>Do not display the default error message to the user and continue.
<dt><code>acDataErrDisplay</code>
<dd>Display the default error message to the user.  This is the default value
	of the <code>Response</code> parameter.
</dl>

</dl>

<p>Here is the list of event procedures which use <code>Cancel</code>
and/or <code>Response</code> parameters.

<dl>
<dt><code>Sub Form_Open (Cancel As Integer)</code>
<dt><code>Sub Form_Unload (Cancel As Integer)</code>
<dt><code>Sub Form_BeforeInsert (Cancel As Integer)</code>
<dt><code>Sub Form_BeforeUpdate (Cancel As Integer)</code>
<dt><code>Sub Form_Delete (Cancel As Integer)</code>
<dt><code>Sub Form_BeforeDelConfirm (Cancel As Integer, Response As Integer)</code>
</dl>

<p>There are more event procedures with other parameters. One of them is:

<dl>
<dt><code>Sub Form_AfterDelConfirm (Status As Integer)</code>

<dd>
<p>The <i>After Del Confirm</i> event cannot be canceled, but you can check the status
of the deletion.  <code>status</code> is the output argument that is assigned
one of the following values:</p>

<dl>
<dt><code>acDeleteOK</code>
<dd>The record has been deleted.
<dt><code>acDeleteCancel</code>
<dd>The deletion has been canceled by Visual Basic.
<dt><code>acDeleteUserCancel</code>
<dd>The deletion has been canceled by the user.
</dl>

<p>Here is another task for you.</p>

<table><tr><td class="notec">
Create event procedures for all events caused by the deletion of a record. 
Let the user be able to cancel the operation.  Notify him/her 
of everything that happens. 
</table> 

<dt><code>Sub Form_KeyPress (KeyAscii As Integer)</code>
<dd>The user pressed a key. The <code>KeyAscii</code> parameter passes the ASCII code
	of this key.
<dt><code>Sub Form_Error(DataErr As Integer, Response As Integer)</code>
<dd>It allows handling errors which occur in the class of a form.
The <code>DataErr</code> parameter passes the code of this error.
</dl>

<h4>Example 2</h4>

<p>You can cancel events which are handled by procedures with 
the <code>Cancel</code> parameter.
We present an example where the event procedure cancels the event by setting
the parameter <code>Cancel</code> to <code>True</code>.</p>

<p>In the code below, when the <i>Departments</i> form is about to be closed 
(unloaded), the event procedure of this form checks whether the associated
<i>Employees</i> form is opened. If it is, the event is canceled with an
an appropriate message and the <i>Departments</i> form stays open.</p>

<pre>Private Sub Form_Unload (Cancel As Integer)
  If ISOpen("Employees") Then 
    Cancel = True
    MsgBox "Close form Employees first."
  End If 
End Sub </pre>

<hr><h3><a name="Uzycie">Using SQL</a></h3>

<p>Sometimes an SQL statetement (a functional query) should be run as 
a response
to the user's actions on the form.  Such a statement performs changes of the data
stored in the database.  These changes depend on the values of fields of the form.</p>

<p>To run an SQL statement, use the <code>RunSQL</code> method of the 
<code>DoCmd</code> object. Its only parameter is the text of the statement to be executed.</p>

<h4>Example 3</h4>

<p>Imagine the <i>Departments</i> form which displays departments of a
company.  If you want to delete a department with employees, the standard
deletion performed by the form is not sufficient because the employees
cannot be assigned to a non-existing department.  The referential integrity
dissallows doing this.  One of the possible solutions is to assign
<code>Null</code> to the <code>Deptno</code> column in the records of employees
of the deleted department.  We are going to code it in the event procedure <i>On
Delete</i> by means ofthe <code>RunSQL</code> method.</p>

<pre>Private Sub Form_Delete(Cancel As Integer)
  DoCmd.RunSQL "UPDATE Employees " &amp; _
               "SET Deptno = Null" &amp; _
               "WHERE Deptno = " &amp; Me![Deptno]
End Sub</pre>

<p>Pay particular attention to the way the value of the <code>Deptno</code>
field is
appended to the UPDATE statement. If <code>Deptno</code> were a text, we would
surround it by quotes inside the text of the statement.  Therefore, we would write:

<pre>"WHERE Deptno = '" &amp; Me![Deptno] &amp; "'"</pre>

<p>If you need to nest quotes, you can use two kinds of these.  You can nest quotes, if they
are of different kinds.</p>

<p>Furthermore, you cannot use phrases like <code>Form![Employees]![Last Name]</code>
in SQL statements. For example the following call is incorrect.</p>

<pre>DoCmd.OpenForm "Employees", , , "Deptno=Forms![Departments]![Deptno]"</pre>

<h4>Example 4</h4>

<p>A user may fill the combo box with a value of a foreign key that does not
occur in the associated primary key (and does not belong to the row source as well).
If he/she wants to add this value to the referenced table, we can use 
INSERT statement in the event procedure for <i>Not In List</i>.</p>

<p>Let us assume that the user adds a new employee and uses <code>Deptno</code>
of a non-existing department.  The following event procedure will add this 
<code>Deptno</code> to the <i>Departments</i> table.</p>

<pre>Private Sub Deptno_NotInList(NewData As String, Response As Integer)
  Dim ctl As Control
  Set ctl = Me![Deptno]
  If MsgBox("Add new department?", vbYesNo, "WARNING") = vbYes Then
    DoCmd.RunSQL "INSERT INTO Departments (Name) " &amp; _
                                   VALUES (" &amp; NewData &amp; ")"
    Response = acDataErrAdded
  Else
    Response = acDataErrContinue
    ctl.Undo
  End If
End Sub</pre>

<p>Note that we do not fill the <code>Deptno</code> column.  Its data type is
<code>Autonumber</code>.  Therefore this column is automatically
assigned by MS Access.</p>

<p>Now it is time for an exercise.</p>

<p align="center">
<table><tr><td class="notec">

<p>Add a column <i>Car brand</i> to the <i>Employees</i> table.  It will be the
foreign key which references the <i>Car brands</i> table (create that as well). 
Place a combo box <i>Car brand</i> onto the <i>Employees</i> form. Its row source
will be the <i>Car brands</i> table.</p>

<p>Create code which will add a new car brand to the <i>Car brands</i> table
when the user
types in a new brand to the combo box.</p>

<a href="javascript:popUp('ok01.html',700,250)">Which</a>
events should be handled? Write appropriate event procedures. 
</table> 

<hr><h3><a name="Przyk">Multi Select</a></h3>

<p>A list box may allow selecting multiple items. Such list boxes 
must have property
<i>Multi Select</i> set to <i>Simple</i> or <i>Extended</i>.  If this property equals
<i>None</i>, the user can select at most one item.</p>

<p>If we have a list box which allows multiple selection, we can select a
number of items and perform an operation on all of them. If the <i>Multi
Select</i> property is set to <i>Simple</i>, the user can use a space bar
and/or a mouse.  If it is <i>Extended</i>, the user can use keys CTRL and
SHIFT with the mouse buttons.</p>

<p>Let us consider the <i>Departments</i> form which shows departments and
the <i>Employees</i> list box with all employees of every department.  This list
box consists of two columns <i>Empno</i> and <i>Last Name</i>.  <i>Empno</i>
is hidden (its width is 0cm).</p>

<p>After we select a set of employees, we can give a raise to every one of them.</p>

<p align="center"><img border="0" src="https://gakko.pjwstk.edu.pl/materialy/2398/lec/wyklad11/images/11_3.png"></p>

<p>In order to access the selected items of the list box, we scan its
the <code>ItemsSelected</code> property which returns the collection of them and
function <code>Column</code> that returns the value of the item of this
collection.  The first argument of <code>Column</code> is the column number,
while its second argument is the position number in this collection.  The
data type of items of the <code>ItemsSelected</code> collection is
<code>Variant</code>.  In the case presented by the picture the collection
consists of numbers 0, 1, 3, 6. If the selection were single, you could omit
the second argument.</p>

<p>In order to give the raise, we use the <code>RunSQL</code> method of the object
<code>DoCmd</code>. Here is the body of this procedure.</p>

<pre>Private Sub Raise_Click()
  On Error GoTo Raise_Error

  Dim listBox As Control, pos As Variant, userAns As Integer
  Dim Question As String, SQLQuery As String

  Set listBox = Me![Employees]

  For Each pos In listBox.ItemsSelected
    'Values of subsequent columns of a row are 
    'listBox.Column(0, pos), listBox.Column(1, pos)

    Question = "Really raise the salary of " &amp; _
                      listBox.Column(1, pos) &amp; "?" 
    userAns = MsgBox(Question, vbYesNo, "Raise")
    If userAns= vbYes Then
      SQLQuery = "UPDATE Employees SET Sal = 1.1 * Sal " &amp; _
                 "WHERE Empno = " & listBox.Column(0,pos)
      DoCmd.RunSQL SQLQuery
    End If
  Next pos

  listBox.Requery

Raise_Exit:
   Exit Sub

Raise_Error:
   MsgBox Err.Description
   Resume Raise_Exit

End Sub</pre>

<p>Note that the SQL statement is created by appending the value of
<code>listBox.Column(0,pos)</code> to the text of the WHERE condition.</p>


<hr><h3><a name="Podsumowanie">Summary</a></h3>

<p>Lecture 11 discussed the programming of database applications. We showed
built-in objects of VBA that represent commands,
errors, the application and the debug environment.
We presented programming techniques like error handling, the requery of a combo box,
the cancelation of an event, the execution of an SQL statement and the code for the multiple
selection in a list box.</p>

<hr><h3><a name="Slownik">Dictionary</a></h3>

<dl>
<dt><a href="#DoCmd">DoCmd</a>
<dd>The object that represents commands.

<dt><a href="#Err">Err</a>
<dd>The object that presents information on errors that have occured during
	the execution of the code written in VBA.

<dt><a href="#Instrukcja">MsgBox</a>
<dd>The command that displays the message box.

<dt><a href="#Procedury">parameter of an event procedure</a>
<dd>It allows passing data between the event procedure and the execution
	environment, e.g. the <i>Cancel</i> parameter tells the environment
	whether the event should be canceled; the <i>Response</i> parameter
	is used to pass the instruction how to response to the event. 

<dt><a href="#Uzycie">RunSQL</a>
<dd>The method of object <i>DoCmd</i> that runs SQL statements.

<dt><a href="#Przyk">multi select</a>
<dd>The property of a list box that allows multiple selections.

</dl>

<hr><h3><a name="Zadania">Exercise</a></h3>

<p>Are you using the following features in your final project?</p>

<ol>
<li>Command <i>MsgBox</i>.
<li>Methods of object <i>Cmd</i>.
<li>The <i>Requery</i> method to refresh the content of a combo box after its
	row source was changed.
<li>The "On Not in List" event procedure for a combo box.
<li>The <i>RunSQL</i> method to perform changes in the database in the background.
<li>Event procedures with parameters.
</ol>

<p>Have you coded procedures which handle all possible errors yet?</p>
