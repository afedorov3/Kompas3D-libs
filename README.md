# Библиотеки для CAD системы Kompas3D

В проекте используется:
 * [CGridListCtrlEx](https://www.codeproject.com/Articles/29064/CGridListCtrlEx-Grid-Control-Based-on-CListCtrl)
 * [NumericEdit](https://github.com/mghini/numeric-edit)
 * [CSV Parser](https://github.com/AriaFallah/csv-parser)
 * [CSVWriter](https://github.com/al-eax/CSVWriter)
 * [code from iproute2](https://mirrors.edge.kernel.org/pub/linux/utils/net/iproute2/)
 * [code from strace](https://strace.io/)
 * [code from Notepad2](https://www.flos-freeware.ch/notepad2.html)
 * [CreateDirectoryRecursively() by Cygon](http://blog.nuclex-games.com/2012/06/how-to-create-directories-recursively-with-win32/)
 * [bit manipulation code by invaliddata](https://stackoverflow.com/questions/3142867/finding-bit-positions-in-an-unsigned-32-bit-integer)

## BatchExport
Библиотека пакетного экспорта файлов.  
Подробное описание находится в директории help.

## 3DUtils
Библиотека вспомогательных функций 3D моделирования.  
Возможности:
 * Выделение граней согласно выбранным критериям.
 * Изменение цвета выделенных граней.
 * Точное изменение ориентации.  
   Файл пользовательских пресетов ориентации (orient.csv)
   должен располагаться в директории с файлом библиотеки 3Dutils64.rtw
 * Экспорт/импорт переменных детали или сборки в/из CSV файла.  
   При вызове функции экспорта с модификатором Shift, экспорт производится в буфер обмена в формате CSV.  
   При вызове функции экспорта с модификатором Ctrl, переменные экспортируются в буфер обмена в формате,
   совместимом с полем переменных BatchExport.
