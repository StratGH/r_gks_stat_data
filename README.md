# r_gks_stat_data

Программа для использования региональной статистики Росстата в пакете R

Для отображения каталога используется Api росстата (документации не нашел)

http://www.gks.ru/bgd/regl/[db_path]/

где db_path

- B03_14
- B04_14
- B05_14p
- ...
- B14_14p

Соответсвенно году, начиная с 2003

##Параметры##

?DbName -> название базы данных (на каждый год своя)

?List&Id={id} -> список документов в формате xml, а id - айди документа, родителя по отношению к документам в отображаемом каталоге

id=-1 либо отсутсвие параметра для отображения всего каталога

```
<l0>
<l>
<ImgSrc>Doc.gif</ImgSrc>
<name>Предисловие</name>
<ref>/bgd/regl/B04_14/IssWWW.exe/Stg/d010/i010280r.htm</ref>
</l>
<l>
<ImgSrc>Folder.gif</ImgSrc>
<name>ЗАНЯТОСТЬ И БЕЗРАБОТИЦА</name>
<ref>?7</ref>
</l>
</l0>
```

ref содержит ?[id], если документ узловой (каталог), либо ссылку на автоматически формируемый документ
