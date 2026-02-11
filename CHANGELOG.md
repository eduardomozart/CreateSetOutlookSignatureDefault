# Change Log

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/)
and this project adheres to [Semantic Versioning](http://semver.org/).

## UNRELEASED

### Fixed

- Fixed an issue where the ``FirstFileAreNewer()`` function created an empty local file when the remote image file in ``img src`` was missing.
- Fix undefined variable ``boolMapNetworkDrive`` when running the script from a local disk (e.g., ``C:\Scripts\``).

### Feat

- Add ``/log`` parameter to enable logging to a file.

### Changes

- Renamed variables to follow the Exchange Online flow rule [tokens](https://learn.microsoft.com/en-us/exchange/security-and-compliance/mail-flow-rules/disclaimers-signatures-footers-or-headers).
- Removed ``%%OtherPhone%%`` and ``%%OtherFax%%`` variables since they are arrays, not strings.
- Forces the update of the user's Outlook signature (use ``/noforce`` to disable).

## [1.0.0] - 2021-01-06

- Public release.