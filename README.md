# Orderly Room System

This repository is a record of projects focused on orderly room systems and processes and is currently a work in progress.

## Background

While serving in a position that was not necessarily technical, I aspired to creatively solve some problems by automating a majority of the... *tedious* work that I was doing.

With no prior experience in VBA and limited knowledge of the administrative systems I was trying to navigate, I just started coding and this amazing (monstrous) system was the result. 

It was made to pipe information from one adminstrative process to another and is not as straightforward as it could have been, but I like to think that I did my best and came up with some very creative solutions.  

## Information Flow
![System Scope Graph](https://github.com/daiceman825/orderlyroom-system/blob/main/SystemScopeGraph.png)

At the center of it all is an MS Access database (which I cannot upload here until I sanitize it 28NOV22).
The data contained in that database flows to all the other related documents via queries that are dynamically linked to MS Excel workbooks. The data from these queries is processed and used in those documents to create reports and metrics for unit readiness. 
