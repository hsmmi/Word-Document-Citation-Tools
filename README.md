# Word Document Citation Tools

This repository contains Visual Basic for Applications (VBA) scripts designed to manage citations in Microsoft Word documents. The scripts allow you to automatically link in-text citations to their corresponding references in the bibliography and to remove these links and bookmarks when needed. These scripts are compatible with the IEEE citation format and are particularly useful for macOS users, where Visual Basic Regular Expressions are not supported.

## Scripts Included

1. **LinkCitationsToReferences**: This script automatically creates bookmarks at each reference in the bibliography section and links any matching citations in the text to these bookmarks.
2. **RemoveCitationHyperlinksAndBookmarks**: This script removes hyperlinks and bookmarks that are specifically related to citations.

## How to Use

### Prerequisites

Before you can use these scripts, you must enable Developer mode in Microsoft Word:

- Go to `File` > `Options` > `Customize Ribbon`.
- Check the `Developer` checkbox in the right column and click `OK`.

### LinkCitationsToReferences

This macro creates bookmarks for each reference in the bibliography section that starts with `[` and ends with `]`, and then links in-text citations that match these reference numbers to the bookmarks. This functionality is specifically tailored for documents using the IEEE citation format and is optimized for use on macOS, which lacks support for Visual Basic Regular Expressions.

#### Steps to Run:

1. Open your Word document.
2. Press `Alt + F11` to open the VBA editor.
3. Insert a new module: Right-click on any of the objects in the project explorer > Insert > Module.
4. Copy and paste the `LinkCitationsToReferences` subroutine into the module.
5. Run the macro by pressing `F5` or selecting Run from the menu.

### RemoveCitationHyperlinksAndBookmarks

This macro removes all hyperlinks and bookmarks that have been added by the `LinkCitationsToReferences` macro or that conform to the naming pattern starting with "Ref_". It's useful for cleanup before reapplying links in updated documents, particularly on macOS where regular expressions in VBA are not available.

#### Steps to Run:

1. Ensure you're in the Word document where you previously ran `LinkCitationsToReferences`.
2. Open the VBA editor with `Alt + F11`.
3. Insert a new module if not already created.
4. Copy and paste the `RemoveCitationHyperlinksAndBookmarks` subroutine into the module.
5. Run the macro by pressing `F5` or selecting Run.

## Requirements

- Microsoft Word (any version supporting VBA).
- Basic familiarity with running VBA scripts in Microsoft Word.

## Note

- Always back up your document before running these macros to prevent unintended changes.
- Customize the macro if your citation or reference format differs from the standard IEEE format.

## Support

For any issues or questions regarding the macros, please open an issue in this repository or contact the maintainer.
