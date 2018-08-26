# Change Log

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/)
and this project adheres to [Semantic Versioning](http://semver.org/).

## [0.2.2] Updated Message SubType

### Updates

- Removed the GChatMessageType enum since it's not needed/used
- Updated the GChatMessageSubType enum
- Updated the Message SubType for ChannelJoined, ChannelLeft and CardClicked events to use the GChatMessageSubType enum

## [0.2.1] Updated manifest and build script

### Updates

- Manifest updated to include FunctionsToExport list
- Build script updated to only update PowerShellGet during the Publish task

## [0.2.0] First release to PSGallery

### Added

- Base classes and helper functions
