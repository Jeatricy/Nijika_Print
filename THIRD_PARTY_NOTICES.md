# Third-party notices

Nijika Print uses the following components. Their licenses and notices remain applicable to those components; this file does not replace or relicense them.

| Component | Upstream / licensing information |
| --- | --- |
| Python | https://www.python.org/psf/license/ |
| pywebview | https://github.com/r0x0r/pywebview |
| pythonnet | https://github.com/pythonnet/pythonnet |
| clr-loader | https://github.com/pythonnet/clr-loader |
| pywin32 | https://github.com/mhammond/pywin32 |
| PyMuPDF / MuPDF | https://pymupdf.readthedocs.io/en/latest/about.html#license-and-copyright |
| pypinyin | https://github.com/mozillazg/python-pinyin |
| SumatraPDF 3.6.1 | https://www.sumatrapdfreader.org/ |

## SumatraPDF runtime

The Windows Release embeds the unmodified SumatraPDF 3.6.1 executable as a runtime resource. The executable is not stored in this source repository.

The matching upstream source tag is **3.6.1rel**, verified against the upstream repository:

- Source tree: https://github.com/sumatrapdfreader/sumatrapdf/tree/3.6.1rel
- Corresponding source archive: https://github.com/sumatrapdfreader/sumatrapdf/archive/refs/tags/3.6.1rel.zip

SumatraPDF is distributed under the terms stated in that source tree, including its GPL and third-party notices. The upstream source archive is also provided with this application's Release for access alongside the binary.

## Python package notices

The standalone build includes installed package metadata and license files for its direct Python runtime dependencies. PyMuPDF offers AGPL/commercial licensing options; the community runtime used here retains its applicable license.

Microsoft Edge WebView2 Runtime and Office/WPS are external installed applications and are not bundled as browser or office installations.
