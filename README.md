# ReflexiveQuery
A generic, reflexive wrapper for queries in MS Excel Power Query

## Use Case
In 2023 I had a spreadsheet problem:
* continuously changing source data - new rows and changes to what would ideally be 'fixed' source data, but aren't
* continuous comments & notes in 'additional' columns - multiple users manually entering comments; these are placed in 'additional' columns not originally included in the source data

Naturally, using Data -> Refresh will pull in new source rows and regenerate my table ... but either clear or misalign all entries in the 'additional' columns. Urgh.

As has been described by others, Excel does not have a quick, built-in tool to avoid this data loss on refresh.
It is *definitely* possible, however, as noted by:
* [Imke Feldmann](https://www.thebiccountant.com/2016/02/09/how-to-create-a-load-history-or-load-log-in-power-query-or-power-bi/)
* [Matt Allington](https://exceleratorbi.com.au/self-referencing-tables-power-query/#comment-243435)

Imke's method (which is what Matt refers to) works well. I have used it more than once.
I have to admit that it has been painful to set up and maintain this system, however.
I had a go at writing a 'wrapper' that would do the reflexive-mergy-fun bit. I think I've gotten it down to a single, brief query.

## Usage
This project provides a single query that wraps your existing query and adds/maintains additional columns with manually entered data.

Process:
1. back up your data. I'm a trombone player, not a programmer ... there is no guarantee that this will not cause massive data loss, spill your coffee and make your cat very angry.
2. In Excel, add a new query, paste in the ReflexiveQuery.pq code
3. update/change a few lines to match your data (ThisQueryName, Source, KeyCol, ColumnsAdditionalRequired)
4. save & load-to. Check that the merge is done how you like, else you may have to change the JoinKind.LeftOuter to something else.
6. Get to work.
7. to add further columns, just edit the ColumnsAdditionalRequired value
8. to remove a column, edit ColumnsAdditionalRequired, and also manually remove that column from the output table

## Requirements
* MS Excel - Desktop (at least at 2025), with Power Query
* Enough patience to fiddle with the Power Query Advanced Editor
* Your existing table or source of 'incoming data'
* Your list of 'additional' columns that must be maintained on refresh 

## Files
* ReflexiveQuery.xlsx - working example file. If you've done PQ, understand the problem, and get the idea of self-reflexing queries, just grab this and edit to suit your existing code.
* ReflexiveQuery.pq - the actual code, extracted from file above.

## Long Version
The code has plenty of comments in it. If you're looking for more details, though, this may assist (ie: confuse) you.

Imke and Matt both explain the problem and approach as well as possible. My own summary is:
* I have data coming in, and it's a growing list. Like MS Forms responses.
* I want to write comments attached to that data, in (what looks like) the same table
* I don't want to lose my comments or see them mis-aligned when new data comes in.

So, we need:
* your data loaded into Power Query as something like "SourceData".
* your data's unique key column name. Something like "ID", or "RecordID", or "StudentNumber" or whatever. Depends upon your data, but it must be unique per row.
* a list of the 'additional' columns you require (ColumnsRequired): a couple comments columns, or some manual tagging, or what not. These are for manually-entered data.
* the name you'd like to use for the new, wrapped output (I'm creative: OutputQuery).

Processing:

Steps 3 & 4 are the fun bit. Kudos to Imke for working out that it's possible to merge against the _already existing table_.

1. Check to see if there is already a table on an Excel worksheet with manually entered data. If there _is_ manually entered data, we need to keep that data. Load the whole table into PQ as PreviousOutput.
2. Figure out what columns are in PreviousOutput that _aren't_ already in SourceData. The unique columns are the data that we need to maintain. We put those in a list called ColumnsAdditional.
3. Merge. We take the _current_ SourceData and merge against PreviousOutput on the key column. We join LeftOuter: keep all the current SourceData, merge in only matching rows from PreviousOutput. That means deleting a row in SourceData will irrevocably delete/remove the manually entered data ... perhaps you don't want that - if so, you'll need to edit the code to suit.
4. "Expand". That merge attaches a whole row inside a single column (very Dr Who). We only need our manually-entered data columns, so we take the list of ColumnsAdditional and make sure those get expanded.
5. If our list of "ColumnsRequired" has something that _isn't_ in the table, now ... the user has (hopefully) added a new column for manually-entered data. Add that new column.
6. Sort for cleanliness.
7. Check the data and pray that giving this to other people doesn't cause massive data loss.

## To Do
1. More testing.
2. Code that removes columns that 'look like an additional column' but 'not in the additional columns list'.
3. Do something about the merge: this implementation does an inner left, but that might not work for everyone.
