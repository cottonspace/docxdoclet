# DocxDoclet
Doclet which creates Javadoc as Microsoft Word document.

## Homepage

[http://www.csync.net/blog/pc/docxdoclet/](http://www.csync.net/blog/pc/docxdoclet/)

## How to use as Ant task

```xml
<?xml version="1.0" encoding="UTF-8" standalone="no"?>
<project default="javadoc">
  <target name="javadoc">
    <javadoc access="private" additionalparam="-encoding utf-8" packagenames="package,package,.." sourcepath="path,path,..">
      <doclet name="doclet.docx.DocxDoclet" path="docxdoclet-1.0.jar">
        <param name="-file" value="document.docx" />
        <param name="-font1" value="Normal font name" />
        <param name="-font2" value="Tagged font name" />
        <param name="-title" value="SUBJECT" />
        <param name="-subtitle" value="SUBTITLE" />
        <param name="-version" value="VER 1.0" />
        <param name="-company" value="XXX PROJECT" />
        <param name="-copyright" value="COPYRIGHT" />
      </doclet>
    </javadoc>
  </target>
</project>
```

## Copyright and License
All the source code avaiable in this repository is licensed under the **[GPL, Version 3.0](http://www.gnu.org/licenses)**

This product includes software developed by [The Apache Software Foundation](http://www.apache.org/), under the Apache License 2.0
* Apache POI: Copyright 2003-2015 The Apache Software Foundation. This product includes software developed by
The Apache Software Foundation (http://www.apache.org/).
