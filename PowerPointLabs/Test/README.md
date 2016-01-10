## How to add new test for Functional Test?

0. Create the test slides [here](https://github.com/PowerPointLabs/PowerPointLabs/tree/master/doc/test). It should contain the data that can test the feature and the expected results that can be verified.
1. Create a new test class that extends [BaseFunctionalTest](https://github.com/PowerPointLabs/PowerPointLabs/blob/master/PowerPointLabs/Test/FunctionalTest/BaseFunctionalTest.cs). Override the method `GetTestingSlideName` to return the name of the test slides.
2. Create a new method with attribute `[TestMethod][TestCategory("FT")]`. Execute the feature under test and assert the verification at the end.

***Notes***
* Many helpful utility classes can be found [here](https://github.com/PowerPointLabs/PowerPointLabs/tree/master/PowerPointLabs/Test/Util). If a utility class you need is missing, please create it yourself.
* Many useful PowerPoint operations and PowerPointLabs feature proxies can be found [here (interface)](https://github.com/PowerPointLabs/PowerPointLabs/tree/master/PowerPointLabs/TestInterface) and [here (impl)](https://github.com/PowerPointLabs/PowerPointLabs/tree/master/PowerPointLabs/PowerPointLabs/FunctionalTestInterface.Impl). If some operations or feature proxies you need is missing, please create it yourself.
* `Spy++` is very helpful in testing Winform UI.


## How to add new test for Unit Test?

0. If the class under test doesn't depend on `PowerPoint Object Model`, directly create a unit test against the class and can skip the rest of steps.
1. If the class under test depends on `PowerPoint Object Model`, create the test slides [here](https://github.com/PowerPointLabs/PowerPointLabs/tree/master/doc/test). It should contain the data that can test the feature and the expected results that can be verified.
2. Create a new test class that extends [BaseUnitTest](https://github.com/PowerPointLabs/PowerPointLabs/blob/master/PowerPointLabs/Test/UnitTest/BaseUnitTest.cs). Override the method `GetTestingSlideName` to return the name of the test slides.
3. Create a new method with attribute `[TestMethod][TestCategory("UT")]`. Test the class and assert the verification at the end.
 
***Notes***
* Many helpful utility classes can be found [here](https://github.com/PowerPointLabs/PowerPointLabs/tree/master/PowerPointLabs/Test/Util). If a utility class you need is missing, please create it yourself.
* Useful PowerPoint operations can be found [here](https://github.com/PowerPointLabs/PowerPointLabs/blob/master/PowerPointLabs/Test/Util/UnitTestPpOperations.cs). If some operations you need is missing, please create it yourself.
