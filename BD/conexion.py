#Creating connection with MySQL workbench
import mysql.connector # 1. Lo primero que hacemos es importar el módulo que nos permite conectarnos con MySQL:

def conectarBD(): # 2. Creamos la funcion conectarBD
  ''' 3. Del módulo importado llamamos a la función connect pasando la ubicación de nuestro servidor 
      que es 'localhost', el usuario que por defecto al instalar MySQL se creó el usuario 'root' y la clave 
      de ese usuario que tiene por defecto un string vacío: y la base de datos en mi caso 'farmacia' '''
  cnx=mysql.connector.connect(host='localhost', 
                                  user='root', 
                                  passwd='', 
                                  database='farmacia')
  return cnx #4. Retornamos la conexiòn a la Base de Datos

def list_tables():
  list_table=conectarBD()
  lt=list_table.cursor()
  lt.execute('show tables')
  for table in lt:
    print(table)
  list_table.close()
  return lt



def listarCategorias():
  l=conectarBD()
  c=l.cursor()
  c.execute("select * from categoria")
  for tabla in c: #6. Mediante un for podemos ver todas las categorias existentes de la base de datos 'farmacia':         
      print(tabla)#7. Imprime las tablas de farmacia
  l.close() #8. Cerramos la conexiòn con el Servidor de Mysql
  return l #9. Retornamos la lista de Categorias

#conectarBD()#10. Invocamos las funciòn para mostrar el resultado final
listarCategorias()#10. Invocamos las funciòn para mostrar el resultado final
