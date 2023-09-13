# allied-bill
allied-express bill analysis
## Problem:allied 账单分析
## Steps:
###
1.首先打开allied express账单，使用在线工具网址转换文件格式：将pdf格式文件转换成excel文件。
在线转换文件格式工具网址：https://www.ilovepdf.com/pdf_to_excel
###
2.打开xls文件后，找到初始的日期，挨个复制粘贴，然后对齐上面title具体内容的列，如果有下一页的内容未粘贴到上一页最下面那一行，不需要担心，因为py脚本可以自动帮我们合并到上一行。
###
3.修改完成后，把py脚本中的周数的数字改掉，然后直接运行脚本输出就可以了。
