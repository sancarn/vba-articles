# Why do people use VBA?

## Introduction

Recently, I watched a video by YouTuber [ThePrimeTime](https://www.youtube.com/watch?v=eJ7oQ6cUwAw) which details a dev's frustrations with business culture. Prime is an ex-entrepreneur who currently works in software development at Netflix. His views in this video have been criticised for being jaded by [FAANG](https://www.google.com/search?q=FAANG+companies) business cultures he has worked in. I personally don't feel this way. Although there is some truth to the [developer's (`mataroa`'s) article](https://ludic.mataroa.blog/blog/your-organization-probably-doesnt-want-to-improve-things/), I think it misses the root causes of many issues raised. 

I'm working on a [full response](/articles/general-organisation-doesnt-improve.html) to this article, but in the interim I did want to mention something that was discussed briefly as one of the paragraphs of this article.

> I work on a platform that cost my organization an eye-watering sum of money to produce, over the span of two years, and the engineers responsible for it elected to use spreadsheets to control the infrastructure, so we now have a spreadsheet with 400 separate worksheets that powers but one part of this whole shambling mess.

I'm speculating here, but I'd imagine that the business is using VBA to some capacity to control their 400 worksheet collection. So this begs the question:

>  Why do people use VBA?

## Why do people use VBA

In order to answer this question, we must first look at another question - who actually uses VBA in the first place? In 2021 I ran a poll on [/r/vba](http://reddit.com/r/vba) where I asked redditors why they code in VBA.

![_](./resources/reddit-2021-why-do-you-code-in-vba.png)

From these data, we can clearly see that the majority of people who use VBA do so mainly because they have no other choice. Many organisations run their entire business processes with Excel, and when a little bit of automation is required VBA is usually \#1 on the list.

## The versatility of VBA

In the business I currently work for, in the engineering division, we have access to a variety of technologies (automation platforms):

* OnPrem - PowerShell (No access to `Install-Module`)
* OnPrem - Excel (VBA  / OfficeJS (limited access) / OfficeScripts / PowerQuery)
* OnPrem - PowerBI Desktop
* OnPrem - SAP Analysis for Office
* OnCloud - Power Platform (PowerApps, Power BI, PowerAutomate (non-premium only))
* SandboxedServer - ArcGIS (ArcPy)
* SandboxedServer - MapInfo (MapBasic)
* SandboxedServer - InfoWorks ICM (Ruby)
* SandboxedCloud - ArcGIS Online

We also have a number of databases controlled by IT:

* D1. OnPrem   - Geospatial database <!-- GISSTdb OnPrem -->
* D2. OnCloud  - Geospatial mirror   <!-- GISSTdb OnCloud -->
* D3. OnPrem   - SAP database        <!-- SAP ECC -->
* D4. OnCloud  - SAP BW4HANA partial mirror
* D5. OnPrem   - Telemetry platform  <!-- eSCADA -->
* D6. OnPrem   - Sharepoint          
* D7. OnCloud  - Sharepoint Online
* D8. OnCloud  - EDM Telemetry platform
* D9. OnCloud  - Large mirror database <!-- CDP -->
* D10. OnPrem  - LotusNotes database   <!-- ORM -->
* D11. OnPrem  - IBM BPM database      <!-- STORM -->
* D12. OnPrem  - File System
* D13. OnPrem  - Hydraulic Model Information

`D1`-`D13` databases are summarised in the table below listing what types of data are stored in which systems, the importance of the data stored in each database, and whether the database is essentially a replica of OnPrem information:


Data Type       | D1  | D2  | D3  | D4  | D5  | D6  | D7  | D8  | D9  | D10 | D11 | D12 | D13 |
----------------|-----|-----|-----|-----|-----|-----|-----|-----|-----|-----|-----|-----|-----|
Customer Issues | X   | X   | X   | X   |     |     |     |     | X   |     |     |     |     |
Asset Data      | X   | X   | X   | X   |     | X   | X   |     | X   |     |     | X   | X   |
Telemetry Data  |     |     |     |     | X   |     |     | X   |     |     |     |     |     |
Risk Data       | X   | X   |     |     |     |     | X   |     |     | X   | X   |     | X   |
Financial Data  |     |     | X   | X   |     |     |     |     |     |     |     | X   |     | 
Misc Data       | X   | X   |     |     |     | X   | X   |     |     |     |     | X   | X   |

> Note: `D6` would be C tier if it weren't for the fact we continue to store a business critical spreadsheet on Sharepoint OnPrem for compatibility reasons. See case study 3 for details.

And the data's importance / on cloud replication:

Data Type       | D1  | D2  | D3  | D4  | D5  | D6  | D7  | D8  | D9  | D10 | D11 | D12 | D13 |
----------------|-----|-----|-----|-----|-----|-----|-----|-----|-----|-----|-----|-----|-----|
Data Importance | S   | S   | S   | S   | S   | S/C | A   | B   | S   | S   | S   | S   | A   |
OnCloud Replica | N/A | Yes | N/A | Yes | N/A | N/A | No  | No  | Yes | N/A | N/A | N/A | No  |

> Note: Online replicas are de facto replica's in terms of data's importance, although the reason we need to connect to them is diminished

Finally, let's look at our Automation Platforms and how these link to our Data Platforms. Links shown in the diagram are where the automation platform can access the data from the various data resources:

![_](./resources/who-uses-vba-data-platform-vs-automation-platform.png)

> Note: Some of the links from VBA to OnCloud services are based on my attempts alone. There is no doubt in my mind that VBA can interface with SAP BW4HANA and our other cloud services, I just haven't figured out the authentication requirements and protocols yet

Here's where you might start to see an issue. Looks like the only automation platforms which can connect to all the data sources we need is `VBA` and `Powershell`. `Power BI Desktop` has been introduced in our business but doesn't hit all the platforms which `VBA` does, and even if it did `Power BI` cannot be used for process automation where-as `VBA` can, so what's the point making the switch? Users who do use `Power BI` to target these other datasets usually generate CSVs of this other data and store these in cloud sharepoint system, but what generates those CSVs? `VBA`.

Now, we'd love to use a higher level language in our organisation to handle this business automation. However, **every request** for a high level language to be installed across the team/business e.g. `Python` / `Ruby` / `Node` / `Rust` etc. has been rejected by CyberSecurity in favour of technologies like `PowerAutomate`, `PowerApps` which as you can see above barely touch **any** of the data we need. It is supposedly "Against the technology strategic vision of the company" to allow "end-users" access to high level programming languages. Now even if the data access was there in our business, `PowerPlatform` would still be insufficient to perform the majority of our processes because the algorithms required are so complex that a `PowerAutomate` solutions would become infuriating to maintain and incomprehensible to even IT folks \(e.g. See [projection algorithms](https://www.movable-type.co.uk/scripts/latlong-os-gridref.html#source-code-osgridref)\).

Ultimately the stand-out technologies for us are `Powershell v3` (doesn't even support class syntax and cannot install modules), and `VBA`, purely from a versatility standpoint. As a result of this 'monopoly' on technology I and others have spent hundreds of hours building [open source VBA libraries](https://github.com/sancarn/awesome-vba) which augment `VBA` promoting it to a reasonable language by modern standards.

## The maintenance guarantee of VBA

`D10` and `D11` above are intimately linked. In 2000s many of our systems were built on top of [IBM Lotus Notes](https://en.wikipedia.org/wiki/IBM_Lotus_iNotes) databases. In 2019 Lotus Notes was acquired by HCL, and since then longevity of support has been wavering. Support will officially die in June 2024. As a result, since 2019, technology teams have been trying to migrate many of our systems to new technologies. The business spent an eye watering amount of money developing a system using IBM Business Process manager to supercede one of these Lotus Notes databases. The anticipation was that `D11` would be backfilled with all the data from `D10`, once fully built, and `D10` archived.

It's now 2023:

* We are 8 months away from official support dying.
* Technology teams have thrown away their support contact for IBM BPM.
* There is no replacement in sight for both IBM BPM and Lotus Notes databases.
* IBM BPM solution is poorly maintained
* IBM BPM solution has numerous issues and doesn't function as needed
* Solution has been shoehorned into IBM BPM, despite IBM BPM not being fit for purpose
  * i.e. while IBM BPM does come with a REST API, this REST API is borderline useless to Technology teams and SMEs
* The data from `D10` was never actually transferred to `D11`, meaning the business is now using 2 systems instead of 1.
  * `D11` data model doesn't really support `D10` data either.
* Technology teams don't want to hear about waning support contracts.

SME's use these tools on a daily basis, and ultimately it is SME's who need changes to the system. If SMEs use VBA, they can control and maintain the system as needed. They have a maintenance guarantee, something that should be said for IT systems too, but can't be.

## The control of VBA

In a recent project we are building a new all encompassing IT system to supercede a business critical spreadsheet. This would ultimately demote `D6` to C tier importance. The spec for this system was initially simple - Give us a NodeJS server with a MySQL database. Use React for the UI. Give admins/SMEs (subject matter experts) access to the codebase with access to git for code control. IT and SMEs will collaborate to build the system.

* Technology teams **demanded** that admins/SMEs will not have access to the code.
* Technology teams **demanded** that FrontEnd be built in Microsoft PowerApps, to comply with "Strategic Vision".
* Technology teams **demanded** that BackEnd be built in Microsoft Azure Pipelines, to comply with "Strategic Vision".

Unfortunately, as an admin/SME with more development knowledge than many people in technology, these demands do not sit well with me:

* Technology teams do not understand work in the teams thus do not understand business logic and calculations
  * Thus devs writing business logic is error prone.
* Technology teams has frequently abandoned bespoke technology projects, leaving no resource to maintain and improve the system.
  * Collaboration with SMEs will ensure that at least 1 team maintains resource to maintain system.
* SMEs need to ensure that they have confidence in what is produced.
  * How, without observing that the code doesn't work for all edge cases?
    * Unit tests?
        * Perhaps, but without seeing the code how can we verify these unit tests exist? And are ran frequently*.
* SMEs improve and maintain the existing legacy system, and have unparalleled knowledge of how systems interact.
  * Less knowledge can be shared in upskilling Technology teams where it is required.
* SMEs need to ensure all data is transferred and represented correctly in new system.
  * SMEs unable to do this without backend access.

Ultimately, as long as code stays in VBA it is controlled by the SMEs and the business. Technology teams rarely relinquish control to business teams. SMEs can ensure that software is developed properly in a modular fashion and doesn't end up as a cluster of barely working technologies loosely linked together.

## Conclusion

In conclusion, yes, we (and many others in businesses) do choose to use spreadsheets (and VBA) for many tasks within our organisations, there are many reasons for this, including:

* Poor alternatives provided by IT due to security concerns.
* Poor connectivity of alternatives to source systems, usually because they are still WIP.
* Faults in IT strategy which don't account for certain use-cases.
* Unwillingness to collaborate with SMEs due to security and maintenance concerns.
* Lack of training for users/managers/SMEs in alternative systems.
* Users/SMEs wanting some level of control over the business logic in these systems.
* It's the only viable technology which is available to everyone, as it's part of Office.

This does not mean that we are at all blind to VBA's weaknesses though:

* [Why is VBA most dreaded?](./Why%20is%20VBA%20most%20dreaded.html)
* [What is wrong with VBA?](./Issues%20with%20VBA.html)

There's no doubt in my mind that there are some elements of truth to [mataroa's article](https://ludic.mataroa.blog/blog/your-organization-probably-doesnt-want-to-improve-things/). Sometimes management is poor, but more often than not I believe most people in organisations are trying to do the right thing, and are doing whatever they can with the tools that are available to them.

