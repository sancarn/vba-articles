---
layout: post
title:  "Issues with Microsoft Power Platform"
published: false
authors:
  - "Sancarn"
---

This post is currently being developed! 👀🥷 Please come back later! 😄

## Power Automate

### General

* Uncapable of dealing with advanced formats which may be part of a process. E.G. [This link](https://services.arcgis.com/VTyQ9soqVukalItT/arcgis/rest/services/Empowerment_Zones_and_Enterprise_Communities/FeatureServer/3/query?f=pbf&objectIds=257&outFields=CONTADDR1%2CCONTADDR2%2CCONTCITY%2CCONTDEPT%2CCONTEMAIL%2CCONTFAX%2CCONTFSTNM%2CCONTLSTNM%2CCONTMIDNM%2CCONTNMPRE%2CCONTNMSUF%2CCONTORG%2CCONTPHONE%2CCONTSTATE%2CCONTTITLE%2CCONTZIP%2CCOUNTYFIPS%2CDSITE%2CDSITENAME%2CFIPS%2CFULLNAME%2CLINK%2CNAME%2COBJECTID%2CPERIODA%2CPERIODB%2CPERIODC%2CPERIODD%2CSTATEABBR%2CSTATEFIPS%2CShape__Area%2CShape__Length%2CTRACT%2CTRACTYEAR%2CTYPE%2CURBANRURAL&outSR=102100&returnGeometry=false&spatialRel=esriSpatialRelIntersects&where=1%3D1) automates an ArcGIS server and returns a [ProtocolBuffer](https://protobuf.dev/). Parseing protocol buffers is not an easy task in most languages, but in PowerAutomate? No chance really.

### Performance

* Running js in the browser is faster

### Missing features

* Code editor
* Regex
* Hash/Dictionary creation and usage

