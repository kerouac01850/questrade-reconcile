The QuestradeReconcile macro is python code that uses the Questrade application programming interface (API) to fetch account, position, balance, equity, 30 day activity, and dividend data into a LibreOffice spreadsheet file.

This script is meant to be run infrequently to help provide a dashboard view when, for example, re-balancing a portfolio. It is not meant to track real time market conditions.

![Figure 1: Run the QuestradeReconcile Python Macro](Documentation/RunQuestradeMacro.png?raw=True "Figure 1: Run the QuestradeReconcile Python Macro")

The Questrade platform requires that client applications like the QuestradeReconcile script help protect the integrity of the service from abuse. The service provides a clear expectation of the level of service that the API commits to fulfill. See details here: https://www.questrade.com/api/documentation/ratelimiting.

[Documentation](Documentation/QuestradeMacroDocumentation.pdf?raw=True)
