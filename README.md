# A tool converting markdown to an excel file

## files

- files belonging to a directory 'class' is a program converting a markdown file to a excel file.
- 'README.md' is this file.
- 'login_ui.md' is a sample markdown file with format for this tool.
- 'login_ui.xlsx' is a sample excel file converted by this tool.

## How to write test pecification with markdown and to convert it to an excel file

### How to write

The following example is a format for this coverting tools.

- '#' is a large section.
- '##' is a middle section.
- '###' is a small section.
- '####' is a type of test.
- numbering list is ways of testing.
- list is items of checking
- === is a sign for separate sheets.

===

    # Section 1
    ## Section 2
    ### Section 3
    #### Type

    1. first way
    2. second way
    3. ...
    - first checking item
    - second checking item
    - ...

    ===
    ...

===

### How to convert

You have to execute the js file in an environment with Node.js.
Executing it, you have to set 2 params, read file and written file, like the following example.

    node main.js test.md test.xlsx
