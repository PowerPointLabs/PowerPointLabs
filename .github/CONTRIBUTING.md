# Contributing Guide

Please take a moment to review this document in order to make the contribution
process easy and effective for everyone involved.

Following these guidelines helps to communicate that you respect the time of
the developers managing and developing this open source project. In return,
they should reciprocate that respect in addressing your issue or assessing
patches and features.


## Using the issue tracker

The [issue tracker](https://github.com/PowerPointLabs/PowerPointLabs/issues) is the preferred channel for [bug reports](#bugs),
[features requests](#features) and [submitting pull
requests](#pull-requests), but please respect the following restrictions:

* Please **do not** use the issue tracker for personal support requests (use
  [pptlabs-contributors forum](https://groups.google.com/forum/#!forum/pptlabs-contributors) or [Stack Overflow](http://stackoverflow.com)).

* Please **do not** derail or troll issues. Keep the discussion on topic and
  respect the opinions of others.


<a name="bugs"></a>
## Bug reports

A bug is a _demonstrable problem_ that is caused by the code in the repository.
Good bug reports are extremely helpful - thank you!

Guidelines for bug reports:

1. **Use the GitHub issue search** &mdash; check if the issue has already been
   reported.

2. **Check if the issue has been fixed** &mdash; try to reproduce it using the
   latest `master` in the repository.

3. **Isolate the problem** &mdash; create an example .PPT file that reproduces 
   the bug or take some screenshots.

A good bug report shouldn't leave others needing to chase you up for more
information. Please try to be as detailed as possible in your report. What is
your environment? What steps will reproduce the issue? What version of Office and OS
experience the problem? What would you expect to be the outcome? All these
details will help people to fix any potential bugs.

Example:

> Short and descriptive example bug report title
>
> A summary of the issue and the Office/OS environment in which it occurs. If
> suitable, include the steps required to reproduce the bug.
>
> 1. This is the first step
> 2. This is the second step
> 3. Further steps, etc.
>
> `<url>` - a link to the example PPT slides that reproduce the bug.  
> `<img>` - any screenshots.
>
> Any other information you want to share that is relevant to the issue being
> reported. This might include the lines of code that you have identified as
> causing the bug, and potential solutions (and your opinions on their
> merits).


<a name="features"></a>
## Feature requests

Feature requests are welcome. But take a moment to find out whether your idea
fits with the scope and aims of the project. It's up to *you* to make a strong
case to convince the project's developers of the merits of this feature. Please
provide as much detail and context as possible.


<a name="pull-requests"></a>
## Pull requests

Good pull requests - patches, improvements, new features - are a fantastic
help. They should remain focused on the selected issues and avoid containing unrelated commits.

For your first pull request, select an issue labelled `forFirstTimers`. For subsequent pull requests, prefer those labelled `forContributors` and with higher priority.

**Please reference the selected issue** like this `#{issue-number}` in the pull request. When the pull request is ready for review, apply label `status.toReview` to it.

**Please ask first** before embarking on any significant pull request (e.g.
implementing features, refactoring code, porting to a different language),
otherwise you risk spending a lot of time working on something that the
project's developers might not want to merge into the project.

**Please ensure** that all pull-requests are tested on PowerPoint 2010/2013/2016 before submitting for approval. Your changes are expected to work on all of these versions.

**Please adhere** to the coding conventions used throughout a project and pass the CI checks (at least not increase the number of errors).

Follow this process if you'd like your work considered for inclusion in the
project:

1. [Fork](http://help.github.com/fork-a-repo/) the project, clone your fork,
   and configure the remotes:

   ```bash
   # Clone your fork of the repo into the current directory
   git clone https://github.com/<your-username>/PowerPointLabs.git
   # Navigate to the newly cloned directory
   cd PowerPointLabs
   # Assign the original repo to a remote called "upstream"
   git remote add upstream https://github.com/PowerPointLabs/PowerPointLabs.git
   ```

2. If you cloned a while ago, get the latest changes from upstream:

   ```bash
   git checkout dev-release
   git pull upstream dev-release
   ```

3. Create a new topic branch (off the main project development branch) to
   contain your feature, change, or fix:

   ```bash
   git checkout -b <topic-branch-name>
   ```

4. Commit your changes in logical chunks. Please adhere to these [git commit
   message guidelines](http://tbaggery.com/2008/04/19/a-note-about-git-commit-messages.html)
   or your code is unlikely be merged into the main project. Use Git's
   [interactive rebase](https://help.github.com/articles/interactive-rebase)
   feature to tidy up your commits before making them public.

5. Locally merge (or rebase) the upstream development branch into your topic branch:

   ```bash
   git pull [--rebase] upstream dev-release
   ```

6. Push your topic branch up to your fork:

   ```bash
   git push origin <topic-branch-name>
   ```

7. [Open a Pull Request](https://help.github.com/articles/using-pull-requests/)
    with a clear title and description.

**IMPORTANT**: By submitting a patch, you agree to allow the project owner to
license your work under the same license as that used by the project.


<a name="dogfooding"></a>
## Dogfooding

We regularly publish the dev version of PowerPointLabs add-in, and a contributor could help to use and verify the updates. The dev version can be downloaded [here](http://www.comp.nus.edu.sg/~pptlabs/download/dev/PowerPointLabs.zip), or it can be built from the recent commit that has a tag. 

If any strange behaviour or exception is encountered, please submit a [bug report](#bugs).

## Coding Standards

See [OSS-Generics Coding Standards](https://github.com/oss-generic/process/blob/master/docs/CodingStandards.adoc)


<a name="branches"></a>
## Branches convention

### Default branches
- `master` holds the RC (release-candidate) version.
- `dev-release` holds the dev-release/develop/dogfooding version with corresponding installer settings.
- `release-standalone` holds the public-release version with corresponding installer settings.
- `release-web` holds the public-release version with corresponding installer settings.

### Feature branches & issue branches
- Feature branch should be named under the feature's name
- Issue branch should be named in this format `{issue number}-issue-short-name`, e.g. `1234-support-abc-def-ghi`.

<a name="release"></a>
## Release strategy

We follow the flows below as our release strategy:
- Normal development flow: 
  - `developers` submit pull request to `dev-release` branch
  - `reviewers` review & merge codes
  - `deployers` do dev-release from `dev-release` branch for dogfooding
  - `owners` sign off RC and merge from `dev-release` to `master` branch
  - `testers` sign off QA and merge from `master` to `release-standalone/web` branch
  - `deployers` do public-release from `release-standalone/web` branch and new codes go LIVE
- Hot-fix development flow: 
  - `developers` submit hot-fix pull request to `master` branch
  - `reviewers` review & merge codes
  - `testers` sign off QA and merge from `master` to `release-standalone/web` branch
  - `deployers` do public-release from `release-standalone/web` branch and hot-fix codes go LIVE

Most text of this document is taken from this [issue-guidelines](https://github.com/necolas/issue-guidelines).
