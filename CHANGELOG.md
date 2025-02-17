# CHANGELOG

This log is intended to keep track of backwards-incompatible changes, including
but not limited to API changes and file location changes.  Minor behavioral
changes may not be included if they are not expected to break existing code.

* Export `NaN` values to `#NUM!` and infinite values to `#DIV/0!`

## v2.6.0

* Stable Version. Fixed Column Width bugs on xStep Parameter
* Added New feature to handle auto filter support on the headers. Just add autoFilterForAllColumn to `true`.
* Reduced the size of the package. Trying to improve more.

## v2.5.0

* It does support bigHeading properly. From v2.0.9 - v2.4.0 was not stable due to continues work on BigHeading Feature, 
* It does have good default Styling even if you don't pass it.


## v2.0.8

* Stable Version. Older all the version from 2.0.0 has bugs. Please do not use.
* Export / Import Fixed.
* All Type Defination has been fixed.
* #5, #6 Issue resolved.

## v2.0.1

* Fixes few bugs. remove publish scripts from package.json.

## v2.0.0

* Full Project Migrated to Typescript.
* Support Full Type Safety now.
* Add Hide Column attribute in dataSet property.

## v1.1.6

* Full Project Migrated to Typescript.
* Support Full Type Safety now.
* Add Hide Column attribute in dataSet property.
* All the imports are now named imports.
* Examples has updated. Will add screenshots.
* Readme is changed a little.

## v1.1.5

* Again Data in DataSet set as any[]. Nothing to done now. If need TS, need fully. Expose TS interface as well. Will do that later.


## v1.1.4

* Removed TS config for MultidataSet prop as any. column has some issue. I will check it later...

## v1.1.3

* Removed File Saver as a dependency.

## v1.1.2

* Excel file generating properly


## v1.1.1

* Minor Error Fixed

## v1.1.0

* Minor Error Fixed


## v1.0.9

* Minor Error Fixed

## v1.0.8

* Types changed.
* I hope to resolve this ts related issue.

## v1.0.7

* Changed the structure.
* I hope to resolve this ts related issue.

## v1.0.6

* Changed Index.d.ts still, finding out issue.
* I hope to resolve this ts related issue.

## v1.0.5

* Remove older babel packages
* Upgraded all bebel packages and config
* Example Import changed

## v1.0.4

* Remove XLSX package
* Added xlsx-js-style package instead of tempa-xlsx 
* xlsx-js-style uses XLSX package under the hood (0.18.5 --older public version)
* Fixed "Prototype Pollution" vulnerability (CVE-2023-30533)
* Lots of API changes

## 1.0.3

* Include local XLSX package (0.19.3) -- Latest
* trying out styles with XLSX package

## 1.0.0

* Initial Build.
* Fix some type definition.
* trying out styles with XLSX package.
* Updrade XLSX package.