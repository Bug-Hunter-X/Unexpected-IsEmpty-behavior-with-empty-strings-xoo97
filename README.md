# VBScript IsEmpty Bug

This repository demonstrates a subtle bug in VBScript related to the `IsEmpty` function and how it interacts with Variant variables containing empty strings.  The bug arises from the fact that `IsEmpty` doesn't always differentiate between an uninitialized variant and a variant explicitly set to an empty string. This difference in interpretation can lead to unexpected program behavior.

## Bug Description
The `bug.vbs` file contains a function that intends to handle cases where a Variant type parameter is either missing or contains an empty string. The `IsEmpty` function is used to detect this; however, this approach fails to reliably distinguish between an empty string and a truly uninitialized variable.

## Solution
The `bugSolution.vbs` file demonstrates a corrected approach. It utilizes the `Len` function in conjunction with a check for the variable being a Variant.  This way, an empty string will explicitly return a 0 length, clearly distinguishing it from an uninitialized or missing variable.

This repository highlights an often overlooked nuance when programming in VBScript and serves as a reminder to test thoroughly and carefully consider data types when working with variants and comparison operations.