# Notices

The initial English reference detector and freshness rules carry forward behavior developed for the Verseform Windows editor by the same project owner. No Bible text is included.

The 66-book, two-character section-code table used to validate DBS response keys is adapted from the Digital Bible Society [BrowserBible verse-detection project](https://github.com/digitalbiblesociety/browserbible-4/tree/master/verse-detection). No detector runtime, DOM transformation, popup, or multilingual recognition data is incorporated. The adapted table is provided under this license:

MIT License

Copyright (c) 2026 John Dyer and the Digital Bible Society

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

Microsoft Office.js is loaded from Microsoft's required hosted CDN at runtime. The shipped runtime package review follows.

## Runtime package review

`bible-tools` 0.2.4 supplies the 66-book canon names and verse-count bounds used by the local detector. Its published npm metadata identifies George Andersen as author and declares the package under the MIT license. The package contains no Bible text and no runtime dependencies. Source: <https://github.com/gla23/bible-tools>.

MIT License

Copyright (c) George Andersen

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

The production dependency review contains only `bible-tools` 0.2.4. Development tools, including Microsoft's transient desktop sideload helper, are not bundled into the task-pane assets or manifest.
