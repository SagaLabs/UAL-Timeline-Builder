<div align="center">

# ☁ UAL-Timeline-Builder (UTB)

![Azure](https://img.shields.io/badge/azure-%230072C6.svg?style=for-the-badge&logo=microsoftazure&logoColor=white)  
[![License](https://img.shields.io/badge/License-GPLv3-blue.svg?longCache=true&style=flat-square)](/LICENSE)  

The tool is intended to help you in your M365 BEC investigations, or prepare the UAL for import to SIEMs like SOF-ELK 📊

*You can either run it locally or use our solution located at [https://ual-timeline-builder.sagalabs.dk/](https://ual-timeline-builder.sagalabs.dk/)*

The tool is developed by [SagaLabs](https://sagalabs.dk), but feel free to contribute or fork as you like.

</div>


## Purpose
**Ever wondered how to approach a new Business Email Compromise (BEC) case?** 

You know that the customer is using M365 and you have exported the Unified Audit Log, but overwhelmed by the amount of data. 😮‍💨

This tool intends to help you in your M365 BEC investigations or prepare the UAL for import to SIEMs. 📊

The tool features simple but effective filter functionality to make the data more readable. A lot of times the investigations in BEC take a lot of time to acquire & prepare data before it can be analyzed effectively in BEC-type cases. This tool will help you to do it.

The tool is not specifically designed for huge datasets, however, it has been tested on quite large environments and runs well. 


### Features
✅ Import standard CSV based Unified Audit Log. 

✅ Conduct in-browser analysis of M365 BEC investigations.

✅ Option to filter on risky operations, making it easier for you to find those nasty new Mailbox-rules. 

💾 All data is processed with the help of in-browser local storage so that means that no data is sent to a server or stored in a database anywhere. 

💾 Features an export function to make the data available in ndjson, suitable for import directly into SOF-ELK or similar. 

## Installation

### Prerequisites 
- Node.js (version 18.x or newer recommended)
- npm
 
For the utility to work we need to export the unified audit log as a CSV. 

The Unified Audit Log can be found on: https://security.microsoft.com/auditlogsearch 
ℹ️ It requires AuditLogsQuery.Read.All permission to do that. 

### Run it locally

```bash
git clone https://github.com/SagaLabs/UAL-Timeline-Builder
cd UAL-Timeline-Builder
npm install 
# To run development
npm run dev 
```
**Open http://localhost:3000 in your browser.**

```bash
# To build
npm run build
npm start
```


## Feature requests

Find a list of features on https://github.com/SagaLabs/UAL-Timeline-Builder/issues