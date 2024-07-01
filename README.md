# Any2PDF Converter

Any2PDF Converter is a powerful Python tool that simplifies the process of converting various file formats to PDF and merging them into a single document.

## Features

- Convert Microsoft Word documents (.doc, .docx) to PDF
- Convert PowerPoint presentations (.ppt, .pptx) to PDF
- Convert images (.png, .jpg, .jpeg) to PDF
- Merge all converted files into a single PDF document
- Automatic cleanup of temporary files

## Requirements

- Python 3.5+ (Tested on Python 3.8 only)
- Aspose.Slides
- Aspose.Words
- img2pdf
- PyPDF2
- python-dotenv

## Installation

1. Clone this repository:
```
git clone https://github.com/ARISTheGod/any2pdf-converter.git
```

2. Navigate to the project directory:
```
cd any2pdf-converter
```

3. Install the required dependencies:
```
pip install -r requirements.txt
```

## Usage

1. Modify the `.env` file in the project root directory with the following content:
```
INPUT_FOLDER=/path/to/input/folder
OUTPUT_FOLDER=/path/to/output/folder
```

2. Place the files you want to convert in the specified input folder.

3. Run the script:
```
python main.py
```

4. The converted and merged PDF will be saved in the specified output folder as `merged_all_files.pdf`.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

&copy; 2024 Aristeidis Alexandridis.

For any part of this work for which the license is applicable, this work is licensed under the [Attribution-NonCommercial-NoDerivatives 4.0 International](http://creativecommons.org/licenses/by-nc-nd/4.0/) license. See LICENSE.CC-BY-NC-ND-4.0.

<a rel="license" href="http://creativecommons.org/licenses/by-nc-nd/4.0/"><img alt="Creative Commons License" style="border-width:0" src="https://i.creativecommons.org/l/by-nc-nd/4.0/88x31.png" /></a>

Any part of this work for which the CC-BY-NC-ND-4.0 license is not applicable is licensed under the [Mozilla Public License 2.0](https://www.mozilla.org/en-US/MPL/2.0/). See LICENSE.MPL-2.0.

Any part of this work that is known to be derived from an existing work is licensed under the license of that existing work. Where such license is known, the license text is included in the LICENSE.ext file, where "ext" indicates the license.

NOTE: As contributor you have no copyright or a License & everything you contributing can be used by us commercially, sorry for that :(

## Disclaimer
```sh
THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS “AS IS” AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING,
BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.
IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY,
OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS;
OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE,
EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY,
WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
```

## Acknowledgments

- Aspose for their document processing libraries
- The open-source community for various Python packages used in this project




