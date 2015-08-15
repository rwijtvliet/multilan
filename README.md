# multilan
Multilanguage documents in MS Word

## Introduction
When dealing with various language-versions of the same document, there are some problems that keep on arising:
-	Adding/removing/changing text in one file needs to be repeated in the others.
-	Changing the layout in one file needs to be repeated in the others.  
In short, work needs to be repeated, and things are bound to be forgotten. 

In an ideal world, you would first make the document in one language, and only when you are absolutely positive that it's the way it will stay, would you translate them. However, this is often impossible, especially when the document is 'alive' and gets various updates throughout its lifetime.

Because of this problem, I want to keep all information (all languages) in one file, and extract single-language versions from it. 
As an additional feature, I want to be able to indicate that some text parts have multiple options. I want to decide which option for each text block should be used each time I create a single-language document.

That is what this project does. It defines rules for creating a 'master document', and uses VBA to turn that master document into a one or more single-language documents.

## Master document
Which language a piece of text in the master document is in, is indicated with special characters, which are the curly brackets { and }. The first character after the opening bracket (the 'tag') is used to characterize the text. 

There are 2 types of tags:

- Lower-case letter. These are used to indicate the language of a text; same letters indicating the same language. I use 'e' for Englisch, 's' for Spanish, 'd' for German. You can pick your own, as long as they are used consistently.  
E.g.: {eName}{sNombre}: Sjaak

- Number. This is used to indicate various 'options', of which only one should end up in the final document. The first must be 1, then 2, etc. There can be 5 options for each given text block (an 'option group'). When starting a new option group, simply start with 1 again.  
E.g.: I'd like to have {1pizza}{2curry}{3a hamburger} for dinner tonight, together with {1coke}{2beer}.

Concerning nesting of tags: languange tags cannot not be nested inside of language tags, and option groups cannot be nested inside of option groups. However, an option group can be nested inside of a language tag, and vice versa. 

## Example file
See `math.docx`.

## Compiling single-language documents
When turning the master document into a single-language document ("compiling"), the following is done to the new document: (the original is left untouched)
*	All text belonging to other languages (i.e., with other language tags) is removed.
*	All comments are removed.
*	You are asked, which option you want to keep in each group.

## Consequences / limitations
*	Because the curly brackets are special characters, these can no longer be used in the text. I could have implemented escape characters, but I didn't at this point.
*	Because highlighting is used as well, you can no longer use highlighting in the text.

## How to run
Keep the document `multilan.docm` open, and also open your master document. Then, with the focus on this document, press ALT-F8 and run the macro 'showForm'. Make sure you have macros enabled. Should run in all MS Office versions >=2007, but I only tested with Office 2010 under Windows 8.1. 
