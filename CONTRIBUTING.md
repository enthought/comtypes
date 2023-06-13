# Contributing to `comtypes`

Welcome! :smile:

`comtypes` is a community project for developing a lightweight COM client and server framework coded by **pure** `Python` on Windows environments.

We appreciate all contributions, from reporting bugs to implementing new features.

## Table of contents
- [To keep the community healthy and sustainable :busts_in_silhouette:](#to-keep-the-community-healthy-and-sustainable-busts_in_silhouette)
- [Reporting bugs :bug:](#reporting-bugs-bug)
- [Breaking troubles down :flashlight:](#breaking-troubles-down-flashlight)
- [Suggesting enhancements :sparkles:](#suggesting-enhancements-sparkles)
- [Contributing to the codebase :open_file_folder:](#contributing-to-the-codebase-open_file_folder)
- [Contributing to documentation :books:](#contributing-to-documentation-books)
- [Final words :green_heart:](#final-words-green_heart)

## To keep the community healthy and sustainable :busts_in_silhouette:

If you have to understand all that this package has to offer, it would require an enormous amount of knowledge.

- `Python`:snake:
- `C`-lang:computer:
- COM interface, implementation, client and server:door:
- COM type library functionalities:wrench:

However, there is **no means to say** that you must understand all of these things to be a contributor.

The purpose of this document is to provide a pathway for you to contribute to this community even if you only know a small portion of this package.

Please keep the followings in your mind:
- :point_right:Please follow the following guidelines when posting a [GitHub Pull Request](https://github.com/enthought/comtypes/pulls) or filing a [GitHub Issue](https://github.com/enthought/comtypes/issues) on this project.
- :bow:The community may not be able to process and reply to your issue or PR right-away. Participants in the community have a lot of work to do besides `comtypes`, but they would try their best.
- :book:For code of conduct, please read [Contributor Covenant](https://www.contributor-covenant.org/).

## Reporting bugs :bug:

We use [GitHub issues](https://github.com/enthought/comtypes/issues) to track bugs and suggested enhancements. You can report a bug by opening a [new issue](https://github.com/enthought/comtypes/issues/new/choose).

We need several infomations for breaking troubles down.

This package handles functionalities of COM libraries in the user's environment.  
Therefore, even if a community participant would like to work on a solution to a problem, the COM library may not be available in their development environments.  
Also, it is possible that you can only reproduce this in your own environment.  
If the community requests you to "try do-something", please respond to them.

At least, please write the following.

### Environment data
- OS: 
- `Python` version: 
- `comtypes` version: 
- what COM type libraries you want to use: 
### Code snippet
- Please provide a minimal, self-contained code snippet that reproduces the issue.
- GitHub permalink is very useful as references.
### Situation reproducing steps
- Please provide a minimal description of conditions and steps needed to reproduce the situation.
### Expected behavior
- The results what you wanted.
### Actual behavior
- The results what you got.

## Suggesting enhancements :sparkles:

We use [GitHub issues](https://github.com/enthought/comtypes/issues) to track bugs and suggested enhancements. You can suggest an enhancement by opening a [new issue](https://github.com/enthought/comtypes/issues/new/choose). Before creating an enhancement suggestion, please check that a similar issue does not already exist.

Please describe the behavior you want and why, and provide examples of how `comtypes` would be used if your feature were added.

## Contributing to the codebase :open_file_folder:

### Picking an issue
Pick an issue by going through the [issue tracker](https://github.com/enthought/comtypes/issues) and finding an issue you would like to work on. Feel free to pick any issue that is not already assigned. We use the [`help wanted` label](https://github.com/enthought/comtypes/issues?q=is%3Aopen+is%3Aissue+label%3A%22help+wanted%22) to indicate issues that are high on our wishlist.

If you are a first time contributor, you might want to look for [issues labeled `good first issue`](https://github.com/enthought/comtypes/issues?q=is%3Aopen+is%3Aissue+label%3A%22good+first+issue%22).  
The `comtypes` code base is quite complex, so starting with a small issue will help you find your way around!

If you would like to take on an issue, please comment on the issue to let others know. You may use the issue to discuss possible solutions.

### Setting up your local environment
Install a version of `Python` supported by `comtypes` into your Windows environment.  
Start by [forking](https://docs.github.com/en/get-started/quickstart/fork-a-repo) the `comtypes` repository, then [clone your forked repository using `git`](https://docs.github.com/en/repositories/creating-and-managing-repositories/cloning-a-repository).  

### Working on developments
Create a new git branch in your local repository, and start coding!

Tests can be run with `python -m unittest discover -v -s ./comtypes/test -t comtypes\test` command.

When `comtypes.client.GetModule` is called, it parses the COM library, generates `.py` files under `.../comtypes/gen/...`, imports and returns `Python` modules.  
Those `.py` files act like ”caches”.

If there are some problems with the developing code base, partial or non-executable modules might be created in `.../comtypes/gen/...`.  
Importing them will cause some error.  
If that happens, you should run `python -m comtypes.clear_cache` to clear those caches.  
The command will delete the entire `.../comtypes/gen` directory.  
Importing `comtypes.gen.client` will restore the directory and `__init__.py` file.

### Pull requests
When you have resolved your issue, open a pull request in the `comtypes` repository.  
Please include the issue number on the PR comment.  
When enough PRs have been accepted to resolve the issue, please close the issue or mention it to the person(s) involved.

## Contributing to documentation :books:

Documents are in the [`.../docs` directory](https://github.com/enthought/comtypes/tree/master/docs).

Follow the same steps as in [the codebase for contributions](#contributing-to-the-codebase-open_file_folder).

## Final words :green_heart:
Thank you very much for your contributions!  
Hope that your contribution will be more profitable for your project and for the work of the community participants.
