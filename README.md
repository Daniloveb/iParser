# IParser
## Overview

Tool for Asset Management.  
Collection and store inventory data.  
Project stopped at 2015. Arhive store.   
>Questions >> https://t.me/daniloveb


## Project structure:
* MS SQL databse iBase
* C# Service iParser
* Console scripts invRM.vbs, CreateKE.vbs
***

## RU
Комплекс для сбора и хранения инвентаризациюнных данных в сети предприятия.
Использовался в проекте С3 при отсутствии SCCM и для обработки компьютеров без сети.

## Краткое описание
На компьютерах предприятия запускался(локально или по сети) файл сценария invRM.vbs, который создавал xml файл с инвентаризационными данными. 
Данный файл копировался на сервер с установленным сервисом iParser для обработки и добавления данных в базу IBase.
Собираемая информация представляла собой массив переменных из основных разделов WMI.
Подробная информация https://github.com/Daniloveb/iParser/tree/master/__Docs

## Установка
* Создаем DB - https://github.com/Daniloveb/iParser/tree/master/__SQL
* Устанавливаем сервис https://github.com/Daniloveb/iParser/tree/master/__Files - msi
* Настраиваем config file сервиса
* Создаем скриптами xml файлы и обрабатываем их с помощью сервиса.
