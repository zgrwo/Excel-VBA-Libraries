# Security Policy

## Supported Versions

| Version | Supported          |
| ------- | ------------------ |
| 2.1.x   | :white_check_mark: |
| 2.0.x   | :white_check_mark: |
| < 2.0.0 | :x:                |

## Reporting a Vulnerability

**Please do not report security vulnerabilities through public GitHub issues.**

Instead, report them privately via GitHub's [Security Advisories](https://github.com/zgrwo/Excel-VBA-Libraries/security/advisories/new) feature, or email the maintainer directly.

### What to include

- A description of the vulnerability
- Steps to reproduce
- Affected module(s) and version(s)
- Any potential impact

### What to expect

- **Acknowledgment**: Within 48 hours
- **Status update**: Within 5 business days
- **Resolution timeline**: Depends on severity — critical issues are prioritized for immediate patching

### Scope

This project is a VBA library running inside Microsoft Excel. Security considerations include:

- **Formula injection**: Malicious inputs that could execute unintended operations
- **File path traversal**: Unsafe file operations in `FileSystemUtils`
- **XML/JSON parsing**: Malformed inputs that could cause denial of service
- **SQL injection**: Unsafe query construction in `SqlUtils`
- **COM object misuse**: Unintended automation of external applications

## Disclosure Policy

We follow coordinated disclosure. Once a fix is released, we will publish a security advisory crediting the reporter (unless anonymity is requested).
