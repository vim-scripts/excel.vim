excel.vim  
=========

A simple vim plugin for displaying data of an excel file.  

![screen shot](http://img5.tuchuang.org/uploads/2013/12/2013-12-09 21:14:57的屏幕截图.png)

##NOTE:
+ It requires your vim with `python` supported.
  
+ Python module [`xlrd`](https://github.com/python-excel/xlrd) is required on your system.   
  To install it you may run `sudo pip install xlrd`  
  
+ For `vim 7.3` and less, it works well for almost all kinds of file formats,   
  ie. `.xls`,`.xlam`,`.xla`,`.xlsb`,`.xlsx`,`.xlsm`,`.xltx`,`.xltm`,`.xlt` etc   

  But for `vim 7.4`, please add the following to your `.vimrc` file:
  ```
  let g:zipPlugin_ext = '*.zip,*.jar,*.xpi,*.ja,*.war,*.ear,*.celzip,*.oxt,*.kmz,*.wsz,*.xap,*.docx,*.docm,*.dotx,*.dotm,*.potx,*.potm,*.ppsx,*.ppsm,*.pptx,*.pptm,*.ppam,*.sldx,*.thmx,*.crtx,*.vdw,*.glox,*.gcsx,*.gqsx'
  ```
  
+ It will parse all sheets by default, displaying them on different tabs.
  
+ You can only __read__ data of a file, any modifications on any tab will __NOT__ do change to the origin excel file.
  
+ Works best on excel files that contain English characters only.  



[![Bitdeli Badge](https://d2weczhvl823v0.cloudfront.net/yakiang/excel.vim/trend.png)](https://bitdeli.com/free "Bitdeli Badge")
