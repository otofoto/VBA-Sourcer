# ANATOMY OF DOC FILES
INTRODUCTION
    1. WHICH MACROVIRUSES?
    2. WHAT IS A DOCFILE
    3. HOW THE WORD DOCUMENT IS STRUCTURED
    4. ABOUT ANTI-VIRUS
    5. DESCRIPTION OF THE EXAMPLE
    LITERATURE
    ATTACHMENT. SOURCE OF DEMONSTRATION UTILITY.

    INTRODUCTION

   This short article is addressed to those who want to know where
specifically, macro viruses live, how to look at them with eyes, and how they
touch with handles.

    1. WHICH MACROVIRUSES?

   Only about macro viruses for MS Word. Macroviruses for MS Access, MS
Excel, etc. are not covered here.
    In addition, we will keep in mind that we are mainly
macro viruses written in Visual Basic for Application (VBA) for
MS Word versions 97, 2000 and XP. Macro viruses for Word 6.0 / 7.0 are no longer
are relevant, but how things are in Word 2003 is not yet very clear.

    2. WHAT IS A DOCFILE

   For details see / 1,2 /, and here in brief.
   It is a .DOC (or .DOT) file with a complex internal
structure called "structured storage" (structured
storage). MS Word can also create documents in "non-native"
formats, for example, in .RTF, but we will not consider this situation.
    The structured storage rules stipulate that:
    - the file is split into many sectors (the default sector size is
512 bytes);
    - sectors are numbered and linked in chains;

                 + ----------------------- +
                 | V
+ --------- + + ---- + ---- + + --------- + + --------- + + ----- ---- + + --------- +
| Title | | 0 | | 1 | | 2 | | 3 | -> | 4 |
+ --------- + + --------- + + ---- + ---- + + --------- + + ----- ---- + + --------- +
                             | ^
                             + ----------------------- +

    - inside the file there is a header describing storage parameters;
    - inside the file there is a FAT table describing which sectors to which
the chains belong;
    - inside the file there are directories describing the name of the chain and its address
first sector.
    Chains form streams, which contain all
useful information. (Actually, there are also lock bytes in which
may store "unnamed" information about which MS Word does not
knows).

    3. HOW THE WORD DOCUMENT IS STRUCTURED

   See details in / 3,4 /, and here in short.
   In general, a document takes up several streams within a file (but
not necessarily all!). Here is a fairly typical thread configuration for
healthy document:

    1Table
    <1> CompObj
    ObjectPool
    WordDocument <- Document text here
    <5> SummaryInformation
    <5> DocumentSummaryInformation

    But for a document containing macros (virus):

    1Table
    Macros
     VBA
      dir
      __SRP_0 <- +
      __SRP_1 | Here exe-code of macros
      __SRP_2 |
      __SRP_3 <- +
      ThisDocument <- Macro s-code and p-code here
      Tralivali <- And here
      Newmacros <- Here too
      _VBA_PROJECT <- Here is a description of the various parts of the project
     PROJECT
     PROJECTwm
    [1] CompObj
    ObjectPool
    WordDocument <- Document text here
    [5] SummaryInformation
    [5] DocumentSummaryInformation

    The same macro is distributed across different threads:
    1) s-code is its text source compressed by the LZSS method;
    2) p-code is a conversion of the source into an intermediate postfix language like

    ; a = 1234

    00A4 Let
    1234 1234
    0027 ->
    0220 number in the list of variable names

    3) exe-code is its code, linked to execute virtual VBA machine, with commands like:

    Push Argument1
    Push Argument2
    [Stack]: = [Stack] + [Stack + 2]

    S-code and its corresponding p-code lie together in the same stream.
    All such streams begin with signature 01 16 01, according to this feature their 
    easy to distinguish.

    4. ABOUT ANTI-VIRUS

   Actually, in 95% of cases the source text of the macro is quite
enough to recognize or suspect a specific virus
infection. The presence of an infection is usually indicated by the presence of methods:

   1) OrganizerCopy;
   2) .Import / .Export;
   3) .InsertLine, .AddFile, etc.

   The remaining 5% falls on cases:

   1) 1% (in fact, even less often) - when the packed source from
the file is deleted (for example, by some crooked antivirus);
    2) 4% - when the virus is polymorphic.

   You can try this: do not decompile p-code and even more so
ex-code from undocumented Microsoft codes (although so
come in "real" antiviruses), but on the contrary, it's good to drive
known VBA grammar in some Yacc and get a translator
s-code to your own internal code. And according to this code
(stripped of spaces, comments, "concatenated" lines, etc.)
recognize virae and even try to emulate some of their pieces.

    5. DESCRIPTION OF THE EXAMPLE

   The DOC file structure can be examined using the DocFile plugin
Browser to FAR Manager: see the object tree, view
the insides of the stream in the form of a hex dump ... But FAR cannot
unpack and show the source text of the macros hiding inside,
and we can / 3 /. As an example, a console utility that
traverses streams in the doc file, finds the signature 01 16 01 in them, finds in
these streams the packed source, unpacks it, looks for
substrings '.Import' and '.Export' and swears if suddenly found. Utility
uses OLE2.DLL library and its methods to work with structured
storage / 5 /. In addition, the utility will not work under Windows 95/98
and will quietly end, because these operating systems only have
stub of the NTDLL.DLL library. Solutions to the problem: 1) rip out
"good" library code for the RtlDecompressBuffer procedure; 2)
to unpack the source codes of macros on your own, it's quite simple.

    LITERATURE

    1. Martin Schwartz. Laola file system.
    2. DrMAD. Internal document format MS Word.
    3. DrMAD. Direct access to macros in MS Word documents.
    4. DrMAD. How to Catch Your Ponytail. Chapter 6. Fight against macro viruses.
    5. Artem Kaev. ActiveX step by step.

   ATTACHMENT. SOURCE OF DEMONSTRATION UTILITY.
  
  Usage: SVIR.EXE docfile
   ![image](https://user-images.githubusercontent.com/26463333/107124856-cd20a900-68ae-11eb-9a20-63e8eb0cbe66.png)
