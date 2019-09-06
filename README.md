The QuestradeReconcile macro is python code that uses the Questrade application programming interface (API) to fetch account, position, balance, equity, 30 day activity, and dividend data into a LibreOffice spreadsheet file.

This script is meant to be run infrequently to help provide a dashboard view when, for example, re-balancing a portfolio. It is not meant to track real time market conditions.

![Figure 1: Run the QuestradeReconcile Python Macro](Documentation/RunQuestradeMacro.png?raw=True "Figure 1: Run the QuestradeReconcile Python Macro")

The Questrade platform requires that client applications like the QuestradeReconcile script help protect the integrity of the service from abuse. The service provides a clear expectation of the level of service that the API commits to fulfill. See details here: https://www.questrade.com/api/documentation/rate-limiting

[PDF Documentation](Documentation/QuestradeMacroDocumentation.pdf?raw=True)

QuestradeReconcile is free software: you can redistribute and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

QuestradeReconcile is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details. You should have received a copy of the GNU General Public License along with this program.  If not, see https://www.gnu.org/licenses/.
