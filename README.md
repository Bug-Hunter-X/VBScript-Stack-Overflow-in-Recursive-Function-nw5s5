# VBScript Stack Overflow Example

This repository demonstrates a common error in VBScript: stack overflow due to excessive recursion.

The `bug.vbs` file contains a recursive function to calculate Fibonacci numbers.  For larger inputs, this function will exceed the default stack limit and cause a runtime error.

The `bugSolution.vbs` file shows a solution using iteration instead of recursion to avoid the stack overflow.