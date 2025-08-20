# Changelog

All notable changes to WinCCSnap will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- Immediate background listener start after installation
- Better error handling during installation
- Enhanced status reporting

### Fixed
- Installation process now works without restart
- Improved PowerShell job management

## [1.0.0] - 2024-08-19

### Added
- Initial release
- PowerShell-based clipboard listener
- Scheduled task for startup
- PNG conversion from CF_BITMAP
- Install/remove/status commands
- Zero-configuration setup
- Battery-friendly settings
- Clean uninstall capability

### Features
- Real-time clipboard monitoring
- Automatic image format conversion
- Background service with minimal CPU usage
- Support for Windows 10+
- PowerShell 5.1+ compatibility