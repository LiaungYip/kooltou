# Kooltou: the MS Outlook Email Backup Tool

![](./doc/img/5.png)

This is a tool for backing up Outlook emails to .msg files.

It has been tested with Outlook 2010.

## Users

* [Download: `./dist/kooltou.exe`](https://github.com/LiaungYip/kooltou/blob/master/dist/kooltou.exe?raw=true) - last updated 2015-07-06
	* The program itself. Download and run this.
	* Does not require installation, therefore does not require elevated rights ("administrator rights").
	* v0.0.3 - 2015-07-06T14:31:47.419000
		* Fix bug where saving to a folder like `F:\Backups\Email` instead tries to save to `F:\`.
		* MD5: `C59C6E4AB72F86CCA6FA86A167BFB3C8  kooltou.exe`
	* v0.0.2 - 2015-07-06T13:04:19.370000
		* Fix bug where 'Saved as MSG' tag was not being applied to emails.
		* Add option for whether to apply `Saved as MSG` tag or not.
		* MD5: `57C640AD3F1893AE8B7BDDE9B69A5E73  kooltou.exe`
	* v0.0.1:
		* MD5: `829278429742826A0699F1DA1EFA9972  kooltou.exe`
* [Documentation: `./doc/README.md`](./doc/README.md)
	* Describes features, instructions for using the software, and troubleshooting tips.

## Developers

* Package requirements:
	* `easygui`
	* `pytz`
	* `pywin32` (`pip` package `pypiwin32`)
	* `unicodedata`
* Build procedure: `pyinstaller --onefile ./kooltou.py`.

## License

Free software under the MIT license. See `LICENSE.txt`.