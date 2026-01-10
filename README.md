# PnP Modern Search - Like component

## Summary

Custom layout for the PnP Modern Search (v4) solution. This layout made for SharePoint pages renders SharePoint pages and news using tags to display basic informations.
This layout also includes a custom web component to perform specific actions on selected pages (promote to news, delete).

![Status-layout](/assets/Status-layout.png)

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.21.1-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [PnP Modern Search (v4)](https://microsoft-search.github.io/pnp-modern-search/)

## Prerequisites

PnP Modern Search (v4) solution must be installed in your tenant.

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| pnp-modern-search-page-status-layout | a1mery |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.0     | January 10, 2026 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run the follwing to build the component:
  - **npm install**
  - **gulp bundle --ship**
  - **gulp package-solution --ship**

- Deploy the component in your tenant ()
- Register the component with your Search Results Web Part: [Register your extensibility library with a Web Part](https://microsoft-search.github.io/pnp-modern-search/extensibility/#register-your-extensibility-library-with-a-web-part)

## Features

This layout renders SharePoint pages and news with a custom web component to delete or promote a page.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
- [PnP Modern Search (v4) - Extensibility possibilities](https://microsoft-search.github.io/pnp-modern-search/extensibility/)
- [PnP/PnPjs](https://pnp.github.io/pnpjs/)
