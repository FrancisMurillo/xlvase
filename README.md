vase
====

Inspired from Python's Nose, this is a small unittesting framework for VBA.

After creating a monolithic psuedo testing module for VBA Excel, I decided I needed to be more organized with how I test and use Debug.Print. It just so happens after using <a href="http://nose.readthedocs.org/en/latest/">Nose</a> in IronPython, there might be some value in replicating a small piece of it.

This executes every module that has the prefix *Test* and every method of with *Test* and gives a summary of the test execution. Simple as that and it comes with its own assertion tools to boot. The phrase here is "*Don't break the vase*", hoping the test suite doesn't break.

**To God I plant these seed, may it turn into a forest**

quick start
-----------

This is a <a href="https://github.com/FrancisMurillo/xlchip">chip</a> project, so you can download this via *Chip.ChipOnFromRepo "Vase"* or if you want to install it via importing module. Just import these four modules in your project.

1. <a href="https://raw.githubusercontent.com/FrancisMurillo/xlvase/master/Modules/Vase.bas">Vase.bas</a>
2. <a href="https://raw.githubusercontent.com/FrancisMurillo/xlvase/master/Modules/VaseLib.bas">VaseLib.bas</a>
3. <a href="https://raw.githubusercontent.com/FrancisMurillo/xlvase/master/Modules/VaseAssert.bas">VaseAssert.bas</a>
4. <a href="https://raw.githubusercontent.com/FrancisMurillo/xlvase/master/Modules/VaseConfig.bas">VaseConfig.bas</a>

And include in your project references the following.

1. **Microsoft Visual Basic for Applications Extensibility 5.3** - Any version would do but it has been tested with version 5.3
2. **Microsoft Scripting Runtime**

So to see if it's working, run in the Intermediate Window or what I call the *terminal*.

```
Vase.RunTests
```

If you done it correctly, you should see some output saying "*The vase is full of air*". This is the main routine to run the test discovery and execution.

Now let's create a sample test module, create a module beginning with the prefix *Test* and copy the following code.

```
Public Sub TestAddition()
    VaseAssert.AssertEqual 1 + 0, 1
End Sub
```

Now if you run the test suite again, you should see output indicating a successful test run.

test discovery
--------------

Like **Nose** and the description above, this executes modules with the prefix *Test* and its methods with the prefix *Test*. 

test execution
--------------

There are some guidelines or things you should remember when creating a test method.

1. **Exception Handling can only go so far** - Since VBA doesn't have an excellent error handling mechanism, I suggest putting an *On Error Goto ErrHandler* or *On Error Resume Next* at the beginning so that execution will not be broken. There are some exceptions you just can't catch.
2. **Assertions do not stop code execution** - So if your creating a very heavy method, make sure it does not depend on the assertion methods to stop the execution like in Java or Python.
3. **Make sure you clean up** - Like with the others, if your opening a database connection or workbook; make sure you create a cleanup routine along with error handling. In my case, I might have created temporary files that weren't deleted. So be wary of that.

You can ignore this guidelines and test will still continue. However the ideal test method template looks like this.

```
Public Sub TestOpenWorkbook()
On Error GoTo Cleanup ' Or Resume Next is a bit more the way I do it
  Dim wb as Workbook
  Set wb = Workbooks.Open(ActiveWorkbook.Path & Application.PathSeparator & "test.xlsx")

  VaseAssert.AssertFalse (wb Is Nothing) ' VaseAssert.AssertIsNothing and VaseAssert.AssertIsNotNothing method?
  
  ' Other assertions you would like to make
Cleanup:
  If Not wb Is Nothing Then
    wb.Close SaveChanges:=False
  End if
  VaseAssert.AssertNotEqual Err.Number, 0, Err.Description ' VaseAssert.AssertErrorNotRaised method?
End Sub
```

solo test execution
-------------------

If you'd like to execute test methods solo, like using F5 with the cursor on it. The assertion methods will give an output as well provided you use *Rewind_* and *Ping_*  at the beginning and at the end. These methods do not affect test suite in general, these are just utility methods to make running test cases one at a time work. 

The former is somewhat optional as it just resets the solo test execution flag; you can remove it but if the output is not the same you might as well put it at the beginning right after the On Error statement. The latter is required to show the output and reset the flag once your done.

You can try it out with this snippet. Run this with F5 or with the Intermediate Window, whichever is faster.

```
Public Sub TestMe()
On Error Resume Next
  VaseAssert.Rewind_ ' Optional but suggested
  
  VaseAssert.AssertTrue True
  VaseAssert.AssertFalse False
  VaseAssert.AssertTrue Falsse
  
  VaseAssert.Ping_ ' Required for solo execution and report
End Sub
```

If you run the above, you can see a similar report like with the actual test suite. This is great if you know which test cases have failed in the suite and want to check each one.

assertion methods
-----------------

These are the currently available *VaseAssert.bas*. All of the methods take an optional parameter called *Message* where if the assertion fails, it shows that message as well. For the most part, the built-in assert failure message should be enough to where the test failed. 

1. **AssertTrue** - Asserts if a condition is True, obviously. This is the go to method for all assertions that aren't presented here.
2. **AssertFalse** - Asserts if a condition is False.
3. **AssertEqual** - An operational asserts if two variants are equal under the equality operator(=). The difference with using this instead of *AssertTrue* is that if there is an error comparing them using the operator such as type mismatch between "1" and True the condition does not raise that error. 
4. **AssertLike** - Another operation assert for the Like operator, this checks if one variable is similar to the other.
5. **AssertGreaterThan/AssertGreaterThanOrEqual/AssertLessThan/AssertLessThanOrEqual/AssertNotEqual** - A set of operational assert signifying the operators >, >=, <, <= and <>. 
6. **AssertInArray** - Asserts if an element is in an array.
7. **AssertArraySize** - Asserts if an zero-based array has the size specified.
8. **AssertEmptyArray* - Asserts if an array is empty or has no element and Empty.
9. **AssertEqualArrasy** - Asserts if two zero-based array has equal size then if it has equal elements.
 

what's next
-----------

There are some things that can be done to improve this.

1. More assertion methods
2. Setup and Teardown for modules
3. Class execution

If you have more ideas, drop me a message.
