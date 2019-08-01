Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports DBaddin.DBAddin

<TestClass()> Public Class UnitTest1

    <TestMethod()> Public Sub TestFunctionSplit()
        Dim check

        check = functionSplit("ignored, because it is before opener..,func(token3,'(', token4,internalfunc(next,next))&this is also ignored, because we have closed the bracket", ",", "'", "func", "(", ")")
        Assert.AreEqual(check(0), "token3")
        Assert.AreEqual(check(1), "'('")
        Assert.AreEqual(check(2), " token4")
        Assert.AreEqual(check(3), "internalfunc(next,next)")
        Assert.AreEqual(UBound(check), 3)

        ' watch out, startStr really searches for the first occurrence ("func") !!
        check = functionSplit("ignoredfunction(because,it,is,before,opener)&func(token3,'(', token4,internalfunc(next,next))&this is also ignored, because we have closed the bracket", ",", "'", "func", "(", ")")
        Assert.AreNotEqual(check(0), "token3")
        Assert.AreNotEqual(check(1), "'('")
        Assert.AreNotEqual(check(2), " token4")
        Assert.AreNotEqual(check(3), "internalfunc(next,next)")
        Assert.AreNotEqual(UBound(check), 3)

        check = functionSplit("ignored(because,it,is,before,opener)&func(token3,'(', token4,internalfunc(next,next))&this is also ignored, because we have closed the bracket", ",", "'", "func", "(", ")")
        Assert.AreEqual(check(0), "token3")
        Assert.AreEqual(check(1), "'('")
        Assert.AreEqual(check(2), " token4")
        Assert.AreEqual(check(3), "internalfunc(next,next)")
        Assert.AreEqual(UBound(check), 3)

        check = functionSplit("func(token3,'(ignore,ignore),whatever is inside'&(still ignored, because in brackets), token4,internalfunc(arg1,anotherFunc(arg1,arg2),arg2))&this is also ignored, because we have closed the bracket", ",", "'", "func", "(", ")")
        Assert.AreEqual(check(0), "token3")
        Assert.AreEqual(check(1), "'(ignore,ignore),whatever is inside'&(still ignored, because in brackets)")
        Assert.AreEqual(check(2), " token4")
        Assert.AreEqual(check(3), "internalfunc(arg1,anotherFunc(arg1,arg2),arg2)")
        Assert.AreEqual(UBound(check), 3)

        ' a different quote and a different delimiter:
        check = functionSplit("=func(token1;token2;""ignoredcloseBracket)""; token3;""ignored1;ignored2"");ignored1;ignored2", ";", """", "func", "(", ")")
        Assert.AreEqual(check(0), "token1")
        Assert.AreEqual(check(1), "token2")
        Assert.AreEqual(check(2), """ignoredcloseBracket)""")
        Assert.AreEqual(check(3), " token3")
        Assert.AreEqual(check(4), """ignored1;ignored2""")
        Assert.AreEqual(UBound(check), 4)
    End Sub

    <TestMethod()> Public Sub TestBalancedString()
        Assert.AreEqual(balancedString("ignored,(start,""ignore '(' , but include"",(go on, the end)),this should (all()) be excluded", "(", ")", """"), "start,""ignore '(' , but include"",(go on, the end)")
        Assert.AreEqual(balancedString("""(ignored"",(start,""ignore '(' , but include"",(go on, the end)),this should (all) be excluded", "(", ")", """"), "start,""ignore '(' , but include"",(go on, the end)")
    End Sub

End Class