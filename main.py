import pymysql
from docxtpl import DocxTemplate


def generate_docx(context):
    tpl = DocxTemplate('template/template.docx')
    tpl.render(context)
    tpl.save('out_generate/generate_%s_doc.docx' % context['name'])


def context_build(table_data, sql):
    context = {'name': table_data[0][0], 'idx': '', 'sql': sql}
    list_ = []
    # 使用 fetchone() 方法获取单条数据.
    for data_ in table_data:
        if data_[2] == 'PRI':
            context['id'] = data_[1]
        list_.append({
            'code': data_[1],
            'is_id': '是' if data_[2] == 'PRI' else '否',
            'date_type': data_[3],
            'is_null': '是' if data_[4] == 'NO' else '否',
            'describe': data_[5],
            'property': data_[5],

        })
    context['list'] = list_
    return context


def connecting():
    # 打开数据库连接
    return pymysql.connect(host='xx',
                           port=3306,
                           user='xx',
                           password='xx',
                           database='xx')


def exec_sql(db, sql):
    # 使用 cursor() 方法创建一个游标对象 cursor
    cursor = db.cursor()
    # 使用 execute()  方法执行 SQL 查询
    cursor.execute(sql)
    # 关闭数据库连接
    return cursor.fetchall()


def review_table(db_name, table_name):
    return '''
        SELECT
    	table_name,
    	column_name,
    	COLUMN_KEY,
    	COLUMN_TYPE,
    	IS_NULLABLE,
    	column_comment 
        FROM
            information_schema.COLUMNS 
        WHERE
            table_schema = '%s' 
            AND table_name = '%s';
        ''' % (db_name, table_name)


def review_table_sql(db_name, table_name):
    return '''
        SHOW CREATE TABLE %s.%s;
        ''' % (db_name, table_name)


if __name__ == "__main__":
    db = connecting()
    table_context_data = context_build(exec_sql(db, review_table('xx', 'xx')),
                                       exec_sql(db, review_table_sql('xx', 'xx'))[0][1])
    db.close()
    generate_docx(table_context_data)
    print(table_context_data)


