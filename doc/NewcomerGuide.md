# Newcomer Guide

## Getting started
Please read this document carefully before you start development on PowerPointLabs. You can always come back and refer when you have doubts/forget about anything.

This document would assume that you have already forked the PowerPointLabs and PowerPointLabs-Website repositories on your local computer.

For official Microsoft documentation on PowerPoint API, visit the website [here](https://msdn.microsoft.com/en-us/library/microsoft.office.interop.powerpoint(v=office.14).aspx).

## Workflow of PowerPointLabs
1. Create a branch from `PowerPointLabs/dev-release`. Your branch name should follow this convention: `issueNumber-description-of-issue` e.g. `123-fix-blur-incorrect-selection`
1. Fix bugs or develop on that branch
1. Do a pull request to `PowerPointLabs/dev-release` from your branch. The pull request title should be same as the issue title with the issue number included (refer to past pull requests). Give a brief description of what you fixed and the reasoning behind your fix to aid the reviewer
1. Most of your codes will be in dev-release until they are properly tested. When ready, they will be merged into master and released to the public

Note: If you are going update any documentation on the PowerPointLabs Github page, create your branch from `PowerPointLabs/master` and make any subsequent pull request(s) to `PowerPointLabs/master` instead of `PowerPointLabs/dev-release`

## Workflow of PowerPointLabs-Website
1. Create a branch from `master`
1. Make desired changes to website and test locally
1. Make a pull request with similar naming convention as above
1. Once the change is accepted, wait for reviewer to merge the branch into `master` and upload the website for the public. Ensure that the [website](http://www.comp.nus.edu.sg/~pptlabs/) can be viewed and displays the latest changes correctly

## Creating Features/Enhancements/New labs
As a rule of thumb:
1. Your codes should refrain from using `#pragma warning disable 0618`, unless absolutely needed. Use methods from `ActionFramework` instead
1. All features/bug fixes should be accompanied by tests and documentation
1. When creating a new lab, use the following existing labs as a guide:
   1. If lab does not require a pane: *Crop Lab*
   1. If lab requires a pane: *Sync Lab*
1. Note that any new labs would require accompanying Unit/Functional Tests. Use existing tests as reference
1. After creating the new lab, you need to create documentation for your features. Documentation should include:
   1. Write-up on how to use the feature: Edit from the current md files for improvement to existing features or start a new md file for completely new features
   1. GIFs (if applicable): Showcase the use of the feature. Keep the GIFs consistent with existing ones and small in size.
   1. Updating of Tutorial file to show users how to use the new features
   1. Please refer to current documentation in the [website](http://www.comp.nus.edu.sg/~pptlabs/docs/)

## Coding Standards Tools
The [Visual Studio Directive Macro](https://github.com/yuhongherald/VisualStudioDirectiveMacro) can help with reorganizing the using directives in your C# files to follow the OSS Generics coding standards. You can refer to the full list here [here](https://github.com/oss-generic/process/blob/master/codingStandards/CodingStandard-CSharp.adoc).

## SonarQube
SonarQube is the tool we are using to monitor the code quality of PowerPointLabs. Some of the issues that SonarQube raises can be ignored, such as `This class has 10 parents which is greater than 5 authorized.`

When fixing bugs in older labs, many errors may pop up from SonarQube. It is up to you if you want to fix them. But we expect minimal problems (if possible, zero) from SonarQube for new code that you write before we merge them.

## User Support
PowerPointLabs uses a google group to handle user support requests. As a developer of PowerPointLabs, it is also your duty to respond to these enquiries. Please respond politely and be understanding of the user’s problems. They may not be as skilled as you in using the computer.

Note that answers to technical questions can be found in the Technical Troubleshooting Guide.
