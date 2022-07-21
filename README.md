
# "POLN" COMPARATOR

Custom easy-to-use comparator script to spot differences between 2 order logs in an excel file.

Out of my beloved one's necessity to save time and reduce erros when spotting differences
between a client's order log and her company's, I created this simple script that reads an excel file
containing both logs, compares them and appends three new sheets: Non-matching orders, number of orders by date from client's log and number of orders by date from the company's log.





## FAQ

#### How do you determine a matching order?

The following criteria must be met:
- "POLN"  and "VN ID" ("IfgNo") values must be the same across logs.
- "Balance" and "Qty Open" values must be the same.
- "Requested Date" and "FechaRequerida" must be the same.
- There must be at least one entry per "POLN" and "VN ID" ("IfgNo") combination.

#### Why do the fields have such weird and inconsistent names between logs?

Prior to this script, the comparison was made manually. The old way to compare was to merge both logs into one file,
keeping the names from each log intact, then apply some excel functions to try and detect as many differences as possible.

The functions implemented where not polished and maybe where not the right ones, so it was still
necessary to check a good bunch of orders manually to solve possible mistakes.

My script takes this file as a starting point, taking existing names and columns, thus reducing the
time required to make every column name match. The purpose of the script is to reduce time.

#### Why not do it with excel?

I'm pretty sure this can be done with excel, but I don't master excel and neither does my GF.
I already know python, pandas and can implement a quick algorithm faster than I can learn excel.

#### But it only takes 5 mins to implement it in excel...

That was never the point. I wanted to do it with python and so I did.

#### What's the company's name? 

Not relevant.

#### Did they pay you?

No. I did this for my GF and as a quick personal project.






## License

[MIT](https://choosealicense.com/licenses/mit/)




