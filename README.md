# JdbcTool
This is a JDBC command line query tool that supposed XLS, HTML, CSV and TEXT output.

This code was inspired by and is based on the JdbcTool by michael@quuxo.com.  That tool can be found at http://quuxo.com/products/jdbctool/

I've **significantly** expanded it, and added some cool stuff so that you can output HTML, XLS, CSV, and TEXT files.  Furthermore, HTML files can be styled using a built-in CSS file, or by supplying your own.  The enhancements were done for me to support emailed HTML attachments, so the HTML format merges the style (css) file into the HTML output.  The XLS mode will allow for "appends" to existing files, so you can have an XLS file with 4 tabs, where each tab is created using a query.

I have also converted this to a maven project.  When you build the code, it will include the following drivers:
  - jTDS
  - Postgres
  - MySql
  - Snowflake (massively distributed cloud database -- https://www.snowflake.net/
