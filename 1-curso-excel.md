---
marp: true
theme: uncover
class: invert
paginate: true
header: 'Curso de Automatización de Datos en Excel'
footer: 'Automatacore.com'
---

# 🚀 Automatización de Datos en Excel 🛠️
---
# 🎉 
# ¡Bienvenido(a)! 
---
# 🧠📚 
# EN ESTE CURSO APRENDEREMOS
---
# 🧹 
# ALMACENAMIENTO 
# EFICIENTE
---

# 📊 
# EN TABLAS 
---

# USANDO MACROS 
##  DE EXCEL 
#  🤖 

---

##  APRENDEREMOS CÓMO USAR MACROS VBA PARA ALMACENAR DATOS DE UN FORMULARIO A UNA TABLA EN EXCEL 
---

PRACTICAREMOS DIVERSOS CASOS HASTA FAMILIARIZARNOS CON EL CÓDIGO DE REFERENCIA Y APLICARLO EN DISTINTOS CONTEXTOS 
# 😃

---

ESTOS CONOCIMIENTOS TE PERMITIRÁN OPTIMIZAR TUS FLUJOS DE TRABAJO Y MEJORAR LA GESTIÓN DE DATOS 
# 📈

---

ES UNA PIEZA CLAVE PARA CONSTRUIR UNA GRAN VARIEDAD DE SISTEMAS FUNCIONALES 
# 🧩.

---

# TE RECOMENDAMOS TENER ALGUNAS BASES ANTES DE COMENZAR
---

# - *Conocimientos Básicos de Excel y...* 📊

---

# - *Nociones de Macros* 🤖

---

# AHORA SÍ, 
# ¡¡COMENZEMOS!! 🚀
---

# ¿Qué son las Macros❓
---

## Las macros son secuencias de instrucciones o comandos que permiten automatizar tareas repetitivas. 

---

# Al crear una macro, se graba un conjunto de acciones realizadas... 

---

y luego estas acciones pueden reproducirse automáticamente con un solo clic o mediante un atajo de teclado. 
# ⚙️

---

# En el contexto de Excel
---

Una macro es un programa pequeño compuesto por código de programación en el lenguaje Visual Basic for Applications (VBA).

---

## VBA permite interactuar con Excel y manipular datos, celdas, hojas y otras funciones del programa.
 # 📋

---

# Resumen
---

Un macro en EXCEL puede memorizar las acciones y también materializar acciones escritas en su propio lenguaje
 #  📝

---

# Entonces tenemos que aprender la forma de comunicarnos con Excel
 #  📚

---

Pero por cuestiones prácticas, omitiremos el proceso de aprender todo el lenguaje de programación VBA.
#  ⏳

---
# 💡
Y Por conveniencia, escogeremos solo un fragmento de código y lo desmontaremos hasta que cada palabra quede clara

---
### Este es el codigo


```vb
Sub AGREGAR()
    With Hoja1.ListObjects("tabla").ListRows.Add
        .Range(1) = Range("celda").Value
    End With
End Sub
```
---

# 📚 Parece simple pero aún tiene mucho que enseñarnos 🧠

---
# Antes de aplicarnos, hagamos un ejemplo de un macro básico 🛠️

---
## ¡Hagamos un botón que al darle clic muestre el mensaje "Bienvenido"! 🖱️

---
# ¡¡AHORA VAMOS A LA PRÁCTICA!! 💪

---
# Perfecto, ¿Ahora qué tal un ejercicio más? 🤔

---
### Realizar un botón que al presionar muestre el valor de la suma de dos datos en un cuadro de mensaje ➕📊

---
Perfecto, ahora en caso que necesites profundizar en macros, puedes revisar los fundamentos. Hay una gran cantidad de material gratuito disponible 📚

---
## Entonces, ¿aprender a codificar es lo único necesario? 🤔
---
No, necesitamos el criterio y la experiencia necesarias para poder aplicarlo en un sistema real y usable... 
# 👩‍💼

--- 
Pero al ser la ruta demasiado extensa, se requiere tiempo. Nosotros vamos a tomar un desvío y subirnos a **hombros de gigantes** 🏞️

---
Por eso recomendaremos un concepto del campo de diseño de software y lo usaremos a nuestra disposición 🤓

---

# Patrón MVC 🏗️

---

El patrón Modelo-Vista-Controlador (MVC) es una arquitectura de diseño que separa la lógica de una aplicación en tres componentes principales:

---

# 🧩
# 1. El Modelo, 
---
# 👁️
# 2. La Vista y 
---
# 🎮
# 3. El Controlador. 
---
Esta separación permite una mejor organización del código y facilita el mantenimiento y la escalabilidad de la aplicación. 🏛️

---
## MASOMENOS ASI

---
```
┌────────┐   Comunican   ┌───────────┐      DATOS      ┌───────┐
│ Modelo ◄──────────────►│Controlador├────────────────►│ Vista │
└────────┘     Datos     └─────▲─────┘                 └───┬───┘
                               │                           │
                               │Solicita                   │
                           ┌───┴───┐      Responde         │
                           │USUARIO│◄──────────────────────┘
                           └───────┘      Visualmente

```
---
# En el caso de Excel 📊
---

La implementación del patrón MVC en Excel mediante macros y VBA permite una mejor organización del código, 🛠️

---

* Facilita el mantenimiento y la escalabilidad de las soluciones automatizadas. 🔄

---

* Con el MVC, se separa la lógica de negocio de la interfaz de usuario y la gestión de datos, facilitando el manejo de aplicaciones complejas. 🤯

---


```
┌────────┐   Comunican   ┌───────────┐      DATOS      ┌───────┐
│ Modelo ◄──────────────►│Controlador├────────────────►│ Vista │
│   de   │     Datos     │  Logica   │                 │       │
│  Datos │               │     ~     │                 │ Inter-│
│   ~    │               │  Codigo   │                 │  faz  │
│ Tablas │               │    VBA    │                 │   ~   │
│   ~    │               │     ~     │                 │Usuario│
└────────┘               └─────▲─────┘                 └───┬───┘
                               │                           │
                               │Solicita                   │
                               │                           │
                           ┌───┴───┐      Responde         │
                           │USUARIO◄───────────────────────┘
                           └───────┘      Visualmente
```
---
## 🧩 Modelo

---

El Modelo representa la lógica de negocio y los datos de la aplicación.

---

Es responsable de gestionar la información y proporcionar los métodos para acceder y modificar los datos.

---

En el contexto de Excel, el Modelo se refiere a las hojas de cálculo, tablas y cualquier otra estructura de datos utilizada para almacenar información.

---

## 👁️ Vista
---

La Vista es la interfaz gráfica con la que el usuario interactúa. 

---
Es responsable de mostrar los datos del Modelo y recopilar la entrada del usuario.

---

En Excel, la Vista serían los formularios o interfaces personalizadas creadas para facilitar la entrada de datos y la interacción del usuario.

---

## 🎮 Controlador

---

El Controlador actúa como intermediario entre el Modelo y la Vista.

---
Es responsable de procesar la entrada del usuario, actualizar el Modelo según sea necesario y actualizar la Vista con los cambios realizados en el Modelo.

---

En Excel, el Controlador correspondería al código VBA que se ejecuta en respuesta a eventos del usuario, como hacer clic en un botón en un formulario.

---
## ¡Listo! Ahora estamos preparados para casos más complejos 🚀

---

# Caso Aplicado 1: Registro de Nombre de Producto 📝
---

**Premisa:** Realizar el envío de un dato llamado "nombre de producto" a una tabla unidimensional llamada "T_Productos" a través de un botón con el macro correspondiente asignado. 📦

---

**Secuencia de desarrollo MVC:**

```
        ┌───────────────┐        
        │Modelo de Datos│        
        └────────┬──────┘        
                 │               
                 v               
 ┌──────────────────────────────┐
 │Diseño de la vista de interfaz│
 └──────────────┬───────────────┘
                │                
                v                
 ┌────────────────────────────┐  
 │Programación del Controlador│  
 └────────────────────────────┘
```


---
Codigo de refencia
```vb
Sub AGREGAR()
    With Hoja1.ListObjects("tabla").ListRows.Add
        .Range(1) = Range("celda").Value
    End With
End Sub
```
---
* Hoja1::Hoja donde esta la tabla Receptora
* Tabla:: Nombre de la Tabla receptora
* Celda::Posicion o  nombre de  celda a enviar

```vb
Sub AGREGAR()
    With Hoja1.ListObjects("tabla").ListRows.Add
        .Range(1) = Range("celda").Value
    End With
End Sub
```
---
```vb
Sub AGREGAR()
    With Hoja1.ListObjects("tabla").ListRows.Add
        .Range(1) = Range("celda").Value
    End With
End Sub
```
```vb
Sub: Función AGREGAR() 
    With:Inicia un bloque para trabajar con la Hoja1 de Excel. 
    .Accede a la objeto (ListObject)  tipo tabla con el nombre 'tabla'.
    .Agregar nueva fila en la tabla a la tabla 

        .Range: Asignar valor de 'celda' a la (1) primera celda de la nueva fila 
    End With: Finaliza el bloque With para la Hoja1 
End Sub:Fin Funcion

```
---
```
                 ┌────────┐                 
                 │Comienzo│                 
                 └───┬────┘                 
                     │                      
                     v                      
             ┌───────────────┐              
             │Proceso Agregar│              
             └───────┬───────┘              
                     │                      
                     v                      
            ┌─────────────────┐             
            │Acceder a la hoja│             
            └─────────┬───────┘             
                      │                     
                      v                     
           ┌────────────────────┐           
           │Accedemos a la tabla│           
           └──────────┬─────────┘           
                      │                     
                      v                     
   ┌────────────────────────────────────┐   
   │Agregamos una nueva fila en la tabla│   
   └──────────────────┬─────────────────┘   
                      │                     
                      v                     
 ┌─────────────────────────────────────────┐
 │Primera columna de la nueva fila <- Celda│
 └────────────────────┬────────────────────┘
                      │                     
                      v                     
       ┌────────────────────────────┐       
       │finalizar acceso a la hoja 1│       
       └─────────────┬──────────────┘       
                     │                      
                     v                      
                   ┌───┐                    
                   │fin│                    
                   └───┘                    

```
---

# Caso Aplicado 2: Registro de Personas

---

Realizar el envío de dos datos: "nombre completo" y "CÓDIGO ÚNICO" a una tabla "T_Personas" con ambos campos respectivos a través de un botón con el macro correspondiente asignado.

---

# Caso Aplicado 3: Registro Corporativo  para  datos no explicitos


---

Realizar el envío de dos datos: "Nombre de la empresa" y "Contacto" a una tabla "T_Empresas" con tres campos: Id, Empresa y Contacto. Siendo Id un dato autoincremental por fórmula. Para enviar la información, usa un botón con el macro correspondiente asignado.

---

# Caso Aplicado 4:  Registro de Acceso y Salida con Múltiples Datos y Controles de formulario

---

Realizar el envio de un formulario a un tabla "T_REGISTRO"  con tres campos: Id, Codigo_persona, Tipo de Registro(ENTRADA/SALIDA), fecha y hora;   Para enviar la informacion usa  boton con el macro correspondiente asignado.


---
# Caso Aplicado 5: Registro de Ventas de un Formulario a dos tablas
---

Realizar el envio de un formulario a dos tablas: (1)"T_VENTAS": ID, Codigo Producto, cantidad, Total; (2)"T_Movimiento_Inventario": ID, Codigo Producto, cantidad,Movimiento(entrada/salida)   Para enviar la informacion usa  boton con el macro correspondiente asignado.

---
![bg left](https://picsum.photos/1080/1080)
``` 
CURSO: Automatización de Datos en Excel
Almacenamiento Eficiente en Tablas usando macros
```
###### Instructor: [Ing. Cesar Uribe](https://www.linkedin.com/in/uribealvites/)
###### Producción: [Ing. Jair Uribe](https://www.linkedin.com/in/jair-uribe/)
###### [Automatacore](https://twitter.com/AutomataCore)



