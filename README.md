`convertTwineProofToPowerPoint` is a tentative project to convert Twine HTML format file to PowerPoint file.

**Usage**
---

```
Convert Twine Proof file to PowerPoint file

Usage:
  convertTwineProofToPowerPoint.py <in_file> <out_file>
  convertTwineProofToPowerPoint.py -h | --help
  convertTwineProofToPowerPoint.py --version

  <in_file>: filename of the Twine file to be converted
  <out_file>: output filename of PowerPoint

Examples:
  convertTwineProofToPowerPoint.py examples/sample.html converted.pptx
    "sample.html" in the folder is converted to a PPTX file
    under the filename of "converted.pptx".

Options:
  -h --help     Show this screen.
  --version     Show version.
```

**Limitations**
---
The program is not exhaustive to cover all the content types in Twine Proof file. This and may miss creating the content or may cause errors in creating PowerPoint file or importing the created PowerPoint to Storyline. 

The expected format of the HTML file includes "tw-passagedata" tags to identify the sections in Twine and to enclose the content in the Twine file.

For example, the Twine file include the following sections.

```
<tw-passagedata pid="1" name="Name of section 1" tags="" position="900,100" size="100,100">
Content
</tw-passagedata>

<tw-passagedata pid="2" name="Name of section 2" tags="" position="900,100" size="100,100">
Content
</tw-passagedata>
```

***Symptom:***

***Cause:***

***Solution:***