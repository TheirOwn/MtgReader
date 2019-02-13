# MtgReader
Search Engine for mtg cards. Credit to mtgjson for the database.

Required:
  Download Allcards.json from MTGJson: https://mtgjson.com/
  Store it in a folder called Cards in the same directory as mtgreader.py

Planning on eventually adding an interface, but that is definitely not my strong suit. 

Syntax: All manacosts are in the original syntax for now: {X} {1} {2}... {W/U/B/R/G} Ex: XGG is {X}{G}{G}
  -Using a "-" sign before a word or phrase marks it as a "banned" word. All results with this word/phrase will be discarded.
  -Using a "~" sign before a word or phrase marks it as a "optional" word. All results with this word/phrase will be marked as better 
   matches.
  

Other: I mostly made this in my free time out of boredom and the desire for a "better" way to search through cards. Namely, Excel. Unfortunately, formatting cards in excel is finicky, so I'll have to keep working to find a layout that is still easily readable.
