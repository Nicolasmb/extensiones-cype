;;-----------------------------------------------------------------------------------------------------------------------------;;
;;  Author:  Nicolas M. Barada, 2022 Growth&Tech Software                                                                      ;;
;;-----------------------------------------------------------------------------------------------------------------------------;;
;;  Version 1.0.0   -   16-04-2022                                                                                             ;;
;;                                                                                                                             ;;
;;  First release.                                                                                                             ;;
;;-----------------------------------------------------------------------------------------------------------------------------;;

;;-----------------------------------------------------------------------------------------------------------------------------;;
;; CRITERIOS PARA NOTACIONES
;;-----------------------------------------------------------------------------------------------------------------------------;;
;;
;; Notacion para objetos, métodos y variables de Autolisp: kebab-case
;; Variables globales entre asteriscos *nombre-variable*
;;
;;-----------------------------------------------------------------------------------------------------------------------------;;


;;-----------------------------------------------------------------------------------------------------------------------------;;
;; FUNCIONES EXCEL
;;-----------------------------------------------------------------------------------------------------------------------------;;
;#region FUNCIONES EXCEL

(defun iniciar-excel ()
  (vl-load-com) ; Se cargan las funciones de ActiveX
  ; Se conecta a una instancia existente de la aplicación o crea una nueva
  (setq *excel* (vlax-get-or-create-object  "Excel.Application"))
  ; Se obtiene una referencia al objeto colección Workbooks, que representa todos los libros abiertos.
  (setq *libros* (vlax-get-property *excel* "Workbooks"))
  ; Para saber la cantidad de archivos (libros) abiertos
  ; (vlax-get-property *libros* "Count")
  ; Se crea un nuevo libro, es un objeto de tipo Workbook.
  (setq *nuevo-libro* (vlax-invoke-method *libros* "Add"))
  ; Se obtiene una referencia a al hoja activa de nuevo libro, es un objeto de tipo Worksheet.
  (setq *hoja-activa* (vlax-get-property *nuevo-libro* "ActiveSheet"))
  ; Se obtiene una referencia a todas las celdas de la hoja activa, es un objeto de tipo Range.
  (setq *celdas* (vlax-get-property *hoja-activa* "Cells"))
)


; ESCRIBIR CELDA
; Función que introduce un valor en una celda de la hoja activa, correspondiente al número de fila y columna 
; que  recibe como argumentos.
(defun escribir-celda (fila col valor)
  (vlax-put-property *celdas* "Item" fila col valor)
)

; CAMBIAR ANCHO COLUMNA
; Función permite cambiar el ancho de una columna, recibe un objeto Worksheet, el indice de una columna, y el ancho deseado
(defun cambiar-ancho-col (hoja col ancho / columnas-obj columna-obj)
  ; Se obtiene una referencia la objeto "Range" que representa a todas las columnas
  (setq columnas-obj (vlax-get-property hoja "Columns"))
  ; Se obtiene una referencia al objeto "Range" que representa la columna por cambiarle el ancho, y se le cambia la propiedad
  ; "ColumnWidth". Esta propiedad espera un valor de tipo Variant.
  (setq columna-obj (vlax-variant-value (vlax-get-property columnas-obj "Item" col)))
  (vlax-put-property columna-obj "ColumnWidth" (vlax-make-variant ancho))
)


; CAMBIAR ALTO FILA
(defun cambiar-alto-fila (hoja fila alto / filas-obj col-obj)
  ; Se obtiene una referencia la objeto "Range" que representa a todas las filas
  (setq filas-obj (vlax-get-property *hoja-activa* "Rows"))
  ; Se obtiene una refernecia al objeto "Range" que representa a la fila por cambiarle la altura, y se le cambia la propiedad
  ; "RowHeight".
  (setq fila-obj (vlax-variant-value (vlax-get-property filas-obj "Item" fila)))
  (vlax-put-property fila-obj "RowHeight" (vlax-make-variant alto))
)


; CENTRAR CONTENIDO CELDA
(defun centrar-celda (fila col / range-obj)
  ; Se obtiene un objeto de tipo "Range" que representa la celda a centrar.
  (setq range-obj (vlax-variant-value (vlax-get-property *celdas* "Item" fila col)))
  ; Se asigna a la propiedad "HorizontalAlignment" un variant con el valor de -4108.
  (vlax-put-property range-obj "HorizontalAlignment" (vlax-make-variant -4108))
  ; Se asigna a la propiedad "VerticalAlignment" un variant con el valor de -4108
  (vlax-put-property range-obj "VerticalAlignment" (vlax-make-variant -4108))
  ; Se ajusta tamaño de la celda al texto.
  (vlax-put-property range-obj "WrapText" :vlax-true)
)


; PONER BORDES A UNA CELDA
(defun put-borders (fila col / range-obj)
  ; Se obtiene un objeto de tipo "Range" que representa la celda a centrar.
  (setq range-obj (vlax-variant-value (vlax-get-property *celdas* "Item" fila col)))
  (vlax-put-property (vlax-get-property range-obj "Borders") "LineStyle" 1)
)


; CAMBIAR COLOR FONDO CELDA
; Recibe el indice de la columna, la fila y un número entero largo que representa un color RGB.
(defun put-background (fila col long-color / range-obj)
  ; Se obtiene un objeto de tipo "Range" que representa la celda a centrar.
  (setq range-obj (vlax-variant-value (vlax-get-property *celdas* "Item" fila col)))
  (vlax-put-property (vlax-get-property range-obj "Interior") "Color" long-color)
)


; TERMINAR EXCEL
(defun terminar-excel ()
  ; (vlax-invoke-method *Excel* 'quit)
  (if (not (vlax-object-released-p *excel*))
    (progn
      (vlax-release-object *nuevo-libro*)
      (vlax-release-object *hoja-activa*)
      (vlax-release-object *libros*)
      (vlax-release-object *excel*)
    )
  )
)

;#endregion


;;-----------------------------------------------------------------------------------------------------------------------------;;
;; FUNCIONES GENERALES
;;-----------------------------------------------------------------------------------------------------------------------------;;
;#region FUNCIONES GENERALES

; SELECCION SET TO LIST
; Convierte un grupo de seleccion en un lista de entidades
; Recibe como argumento un sset.
; Retorna una lista con los nombres de las entidades que pertenecen al sset.
(defun sset-to-list (sset / i lista-entidades)
  (setq i 0)
  (setq lista-entidades nil)
  (while (< i (sslength sset))
    (setq lista-entidades (cons (ssname sset i) lista-entidades))
    (setq i (1+ i))
  )
  lista-entidades
)


; CONDICION
; Función que verifica si una entidad de tipo texto comienza con "%%c" o con "%%C"
(defun cumple-comienza-con-fi (ename / string)
  ; Si el argumento no es nil:
  (if ename
    ; Si la entidad es de tipo texto se procede.
    (if (= (cdr (assoc 0 (entget ename))) "TEXT")
      (progn
        (setq string (cdr (assoc 1 (entget ename))))
        (if (or (= (substr string 1 3) "%%c") (= (substr string 1 3) "%%C")) 
          T
          nil
        )
      )
    )
  )
)


; CONDICION TEXTO CUMPLE PATRON ARMADURAS LOSAS
; Función que verifica si una entidad de tipo texto cumple con un patron dado.
(defun cumple-patron-textos-losas (ename / string)
  ; Si el argumento no es nil:
  (if ename
    ; Si la entidad es de tipo Mtext se procede.
    (if (= (cdr (assoc 0 (entget ename))) "MTEXT")
      (progn
        ; Se obtiene la cadena de texto
        (setq string (cdr (assoc 1 (entget ename))))
        ; Se comprueba si cumple con el patron de las armaduras de losas
        (if (wcmatch string "*#*.%%[C-c]#*L=#*,#*.%%[C-c]#*L=#*")
          T
          nil
        )
      )
    )
  )
)


; LEER CANTIDAD DE BARRAS DE TEXTO ARMADURA DE LOSAS
(defun leer-cantidad-barras-losa (string / sub-string i char)
  (setq i nil)
  (setq sub-string nil)
  ; Se elimina cualquier prefijo
  ; Se va leyendo el primer caracter, si no es de tipo numero de lo quita de la cadena
  ; y se repite el ciclo.
  (while (and (not (wcmatch string "#*")) (< 0 (strlen string)))
    (setq char (substr string 1 1))
    (setq string (vl-string-left-trim char string))
  )
  ; Si existe el caracter divisorio, en este caseo el: " "
  (if (vl-string-search " " string)
    (progn
      ; Se obtiene el indice del caracter divisorio, en este caso el: " "
      (setq i (+ 1 (vl-string-search " " string)))
      ; Se extrae la parte de la cadena de texto a la izquierda del caracter divisorio.
      (setq sub-string (substr string 1 (1- i)))
      ; Se va leyendo el primer caracter, si no es de tipo numero de lo quita de la cadena
      ; y se repite el ciclo.
      (while (and (not (wcmatch sub-string "#*")) (< 0 (strlen sub-string)))
        (setq char (substr sub-string 1 1))
        (setq sub-string (vl-string-left-trim char sub-string))
      )
      ; Si no se encontró el número que representa la cantidad se va a retornar nil.
      (if (= (strlen sub-string) 0)
        (setq sub-string nil)
      )
    )
    ; Sino se encontró un simbolo de diametro en la cadena se va a retornar nil.
    (setq sub-string nil)
  )
  sub-string
)


; LEER DIAMETRO DE BARRAS DE TEXTO ARMADURA DE LOSAS
(defun leer-diametro-barras-losa (string / sub-string i char)
  (setq i nil)
  (setq sub-string nil)
  ; Si existe en la cadena el patron de diametro, en este caso "%%c"
  (if (vl-string-search "%%c" (strcase string t)) 
    (progn
      ; Se obtiene el indice del caracter del simbolo de diametro, %%c o %%C
      (setq i (+ 1 (vl-string-search "%%c" (strcase string t))))
      ; Se extrae la parte de la cadena de texto a la derecha del caracter divisorio.
      (setq sub-string (substr string (+ i 3)))
      ; Se obtiene el indice del caracter divisorio, en este caso el: " "
      (if (vl-string-search " " sub-string)
        (progn
          (setq i (+ 1 (vl-string-search " " sub-string)))
          ; Se extrae la parte de la cadena de texto a la izquierda del caracter divisorio.
          (setq sub-string (substr sub-string 1 (1- i)))
        )
        ; Sino se encontró un simbolo de diametro en la cadena se va a retornar nil.
        (setq sub-string nil)
      )
    )
    ; Sino se encontró un simbolo de diametro en la cadena se va a retornar nil.
    (setq sub-string nil)  
  )
  sub-string
)


; LEER SEPARACION DE BARRAS DE TEXTO ARMADURA DE LOSAS
(defun leer-separacion-barras-losa (string / sub-string i char)
  (setq i nil)
  (setq sub-string nil)
  ; Si existe en la cadena el patron separacion de barras, en este caso "c/"
  (if (vl-string-search "c/" (strcase string t))
    (progn
      ; Se obtiene el indice del patron de separacion de barras
      (setq i (+ 1 (vl-string-search "c/" (strcase string t))))
      ; Se extrae la parte de la cadena de texto a la derecha del caracter divisorio.
      (setq sub-string (substr string (+ i 2)))
      (cond 
        ; Si existe en la cadena el patron "cm"
        ((vl-string-search "cm" (strcase sub-string t))
          (progn
            (setq i (+ 1 (vl-string-search "cm" (strcase sub-string t))))
            ; Se extrae la parte de la cadena de texto a la izquierda del caracter divisorio.
            (setq sub-string (substr sub-string 1 (1- i)))
          )
        )
        ; Si existe en la cadena el patron " "
        ((vl-string-search " " (strcase sub-string t))
          (progn
            (setq i (+ 1 (vl-string-search " " (strcase sub-string t))))
            ; Se extrae la parte de la cadena de texto a la izquierda del caracter divisorio.
            (setq sub-string (substr sub-string 1 (1- i)))
          )
        )
        ; Sino se encontró un simbolo de diametro en la cadena se va a retornar nil.
        (t (setq sub-string nil))
      )
    )
    ; Sino se encontró en la cadena el patron separacion de barras se va a retornar nil.
    (setq sub-string nil)  
  )
)


; LEER LONGITUD DE BARRAS DE TEXTO ARMADURA DE LOSAS
(defun leer-longitud-barras-losa (string / sub-string i char)
  (setq i nil)
  (setq sub-string nil)
  ; Si existe en la cadena el patron separacion de barras, en este caso "c/"
  (if (vl-string-search "L=" (strcase string nil))
    (progn
      ; Se obtiene el indice del patron de separacion de barras
      (setq i (+ 1 (vl-string-search "L=" (strcase string nil))))
      ; Se extrae la parte de la cadena de texto a la derecha del caracter divisorio.
      (setq sub-string (substr string (+ i 2)))
      (cond 
        ; Si existe en la cadena el patron "cm"
        ((vl-string-search "cm" (strcase sub-string t))
          (progn
            (setq i (+ 1 (vl-string-search "cm" (strcase sub-string t))))
            ; Se extrae la parte de la cadena de texto a la izquierda del caracter divisorio.
            (setq sub-string (substr sub-string 1 (1- i)))
          )
        )
        ; Si existe en la cadena el patron " "
        ((vl-string-search " " (strcase sub-string t))
          (progn
            (setq i (+ 1 (vl-string-search " " (strcase sub-string t))))
            ; Se extrae la parte de la cadena de texto a la izquierda del caracter divisorio.
            (setq sub-string (substr sub-string 1 (1- i)))
          )
        )
        ; Si la sub-string se puede convertir a un numero real
        ((numberp (atof (strcase sub-string t)))
	        sub-string
        ) 
        ; Sino se encontró cumple ninguno de las verificaciones de cadena se va a retornar nil.
        (t (setq sub-string nil))
      )
    )
    ; Sino se encontró en la cadena el patron separacion de barras se va a retornar nil.
    (setq sub-string nil)  
  )
)


; FILTRAR TEXTOS ALINEADOS HORIZONTALMENTE CON UNO DE REFERENCIA
; Funcion que filtra todos los textos que estan alineados horizontalmente
; Recibe el texto de referencia y la lista de textos a filtrar.
(defun filtrar-alineados-hor (ename-texto-1 lista-entidades / y-ref tol-alineacion-vert)
  ; Se obtiene la coordenada "y" de la entidad.
  (setq y-ref (nth 2 (assoc 10 (entget ename-texto-1))))
  ; Se remueven de la lista de entidades todas aquellas que no cumplan con la funcion de predicado/test.
  ; La función predicado retorna true para todos aquellos textos cuya diferencia de coordenada "y", respecto a 
  ; la coordenada "y" de "ename-texto-1" sea menor a un determinado limite dado por lim-alineacion-vert.
  ; En este caso lim-alineacion-vert se adoptó como un porcentaje de la altura del texto.
  ; Se obtiene la altura de "ename-texto-1"
  (setq h-texto (cdr (assoc 40 (entget ename-texto-1))))
  ; Se calcula la tolerancia para considerar a dos textos como alineados horizonalmente.
  (setq tol-alineacion-vert (* 0.1 h-texto))
  ; Se remueven de la lista de entidades todas aquellas que no cumplan con la funcion de predicado/test.
  (vl-remove-if-not 
    '(lambda (ename / y-test ent)
       (setq ent (entget ename))
       (setq y-test (nth 2 (assoc 10 ent)))
       (< (abs (- y-ref y-test)) tol-alineacion-vert)
    )
    lista-entidades
  )
)


; FILTRAR TEXTOS A UNA DISTANCIA DETERMINADA DE UN TEXTO DE REFERENCAI
; Funcion que filtra todos los textos que estan a una distancia determinada del texto de referencia (mas una tolerancia)
(defun filtrar-textos-proximos (ename-texto-1 sub-lista-1 / distancia-cumplir distancia tol-distancia)
  ; Se obtiene la coordenada "x" de la entidad.
  (setq x-ref (nth 1 (assoc 10 (entget ename-texto-1))))
  ; Se obtiene la altura de "ename-texto-1"
  (setq h-texto (cdr (assoc 40 (entget ename-texto-1))))
  ; Se establece la distancia que debe haber entre los textos
  ; Esta distancia se obtuvo del cype midiendo la columna "Diam" y la columna "Total (cm)"
  ; La distancia en unidades relativas a la altura del texto de referencia
  ; En este caso la distancia es 27.515 veces la altura del texto ename-texto-1
  ; De esta manera nos independizamos de la escala que podria aplicarse a la tabla
  (setq distancia-cumplir (vlax-ldata-get "dic-configuracion" "distancia-textos"))
  ; Se remueven de la lista de entidades todas aquellas que no cumplan con la función de predicado/test.
  (vl-remove-if-not
    '(lambda (ename / x-test ent dif-x)
       (setq ent (entget ename))
       (setq x-test (nth 1 (assoc 10 ent)))
       (setq distancia (- x-test x-ref))
       ; Se normaliza la distancia
       (setq distancia (/ distancia h-texto))
       ; Se establece la tolerancia para la distancia entre los textos
       (setq tol-distancia 1)
       ; Si la distancia es positiva y es igual a la distancia a cumplir, dentro de una tolerancia
       (and (< 0 distancia) (< (abs (- distancia-cumplir distancia)) tol-distancia))
     )
    sub-lista-1
  )
)


; REINICIAR DICCCIONARIO DE ARMADURAS (diametro . longitud)
; Función que crea reinicia el diccionario de longitudes de armaduras para cada diametro.
(defun dic-armaduras-reset ()
  (vlax-ldata-put "dic-armaduras" "6" 0)
  (vlax-ldata-put "dic-armaduras" "8" 0)
  (vlax-ldata-put "dic-armaduras" "10" 0)
  (vlax-ldata-put "dic-armaduras" "12" 0)
  (vlax-ldata-put "dic-armaduras" "16" 0)
  (vlax-ldata-put "dic-armaduras" "20" 0)
  (vlax-ldata-put "dic-armaduras" "25" 0)
)


; AGREGAR PAR (diametro . longitud) A DICCIONARIO DE ARMADURAS
; Funcion que suma un par (diametro . longitud) en el diccionario dic-armaduras
(defun dic-armaduras-agregar (ename-texto-1 ename-texto-2 / diam longitud longitud-acumulada)
  ; Se obtiene el primer elementos del par, osea, el diametro
  (setq diam (substr (cdr (assoc 1 (entget ename-texto-1))) 4))
  (setq longitud (atof (cdr (assoc 1 (entget ename-texto-2)))))
  ; Se obtiene la longitud actual asociada al diametro, si es nil se le asigna cero.
  (if (null (setq longitud-acumulada (vlax-ldata-get "dic-armaduras" diam)))
    (setq longitud-acumulada 0)
  )
  ; Se suma a la longitud acumulada la longitud.
  (vlax-ldata-put "dic-armaduras" diam (+ longitud-acumulada longitud))
)


; REINICIAR DICCCIONARIO DE LOSAS (diametro . longitud)
(defun dic-losas-reset ()
  (vlax-ldata-put "dic-losas" "6" 0)
  (vlax-ldata-put "dic-losas" "8" 0)
  (vlax-ldata-put "dic-losas" "10" 0)
  (vlax-ldata-put "dic-losas" "12" 0)
  (vlax-ldata-put "dic-losas" "16" 0)
  (vlax-ldata-put "dic-losas" "20" 0)
  (vlax-ldata-put "dic-losas" "25" 0)
)


; AGREGAR PAR (diametro . longitud) A DICCIONARIO DE LOSAS
(defun dic-losas-agregar (cant diam sep long)
  ; Se obtiene la longitud actual asociada al diametro, si es nil se le asigna cero.
  (if (null (setq longitud-acumulada (vlax-ldata-get "dic-losas" diam)))
    (setq longitud-acumulada 0)
  )
  ; Se calcula la longitud total del paquete de barras
  (setq longitud (* (atof cant) (/ (atof long) 100)))
  ; Se suma a la longitud acumulada la longitud.
  (vlax-ldata-put "dic-losas" diam (+ longitud-acumulada longitud))
)


; CAMBIAR COLOR ENTIDAD
(defun cambiar-color-entidad (ename indice-color / ent ename)
  ; Se obtiene la entidad
  (setq ent (entget ename))
  ; Si la entidad ya tiene un color asignado se lo cambia
  (if (not (null (assoc 62 ent)))
    ; Se cambia el color en la lista de asociación y se sobreescribe la variable que 
    ; contiene la entidad
    (setq ent (subst (cons 62 indice-color) (assoc 62 ent) ent))
    ; Si la entidad no tenia color se agrega un par punteado al final de la lista de asociacion
    (setq ent (reverse (cons (cons 62 indice-color) (reverse ent))))
  )
  ; Se modifica la definicion de la entidad
  (entmod ent)
  ; Se actualiza la imagen en pantalla de la entidad
  (entupd ename)
)


; QUITAR COLOR ENTIDAD
(defun quitar-color-entidad (ename / ent layer-name layer-colour)
  (setq ent (entget ename))
  (setq layer-name (cdr (assoc 8 ent)))
  (setq layer-colour (cdr (assoc 62 (tblsearch "LAYER" layer-name))))
  (setq ent (append ent (list (cons 62 layer-colour))))
  (entmod ent)
  (entupd ename)
)


; TRANSFORMAR DICCIONARIO DE ARMADURAS EN TABLA EXCEL
; Recibe como argumento el nombre del diccionario, puede ser:
; "dic-armaduras" o "dic-losas"
(defun dic-armaduras->excel (dic-nombre / data-pares)
  (escribir-celda 1 1 "DIAMETRO")
  (escribir-celda 1 2 "LONGITUD")
  (cambiar-alto-fila *hoja-activa* 1 20)
  (cambiar-ancho-col *hoja-activa* 2 20)
  (centrar-celda 1 1)  
  (centrar-celda 1 2)  
  (put-background 1 1 13431551) 
  (put-background 1 2 13431551)
  (put-borders 1 1)
  (put-borders 1 2)
  (setq i 0)
  (setq data-pares (vlax-ldata-list dic-nombre))
  ; Se ordenas los pares según el diámetro
  (setq data-pares (vl-sort data-pares '(lambda (par1 par2) (< (atoi (car par1)) (atoi (car par2))))))
  (foreach par data-pares 
    (progn
      (escribir-celda (+ i 2) 1 (car par))
      (escribir-celda (+ i 2) 2 (cdr par))
      (centrar-celda (+ i 2) 1)  
      (centrar-celda (+ i 2) 2)
      (cambiar-alto-fila *hoja-activa* (+ i 2) 20)
      (setq i (1+ i))
    )
  )
)


; HIGHLIGHT ENTIDAD
(defun highlight-ent (ename)
  (vla-highlight (vlax-ename->vla-object ename) :vlax-true)
)


(defun unhighlight-ent (ename)
  (vla-highlight (vlax-ename->vla-object ename) :vlax-false)
)
;#endregion


;;-----------------------------------------------------------------------------------------------------------------------------;;
;; FUNCIONES PARA REACTORES
;;-----------------------------------------------------------------------------------------------------------------------------;;
;#region


; CREA LISTA OBJECT WATCHED
; Crea una lista de objetos vlax que serán observados por el reactor "reactor-texto-modificado"
; Recibe una lista de pares de nombres de entidades de tipo texto que representan a los pares diametro y longitud.
(defun crear-lista-object-watched (lista-pares-textos / lista-object-watched)
  ; Se inicializa la lista que contendrá a los objetos vlax que representan los textos de diametros y longitudes.
  ; Estos serán los objetos que serán propietarios (owners) del reactor.
  (setq lista-object-watched nil)
  ; Se toma cada ename de cada par y se lo agrega a la lista de object-watched, previamente transformado a vlax-object
  (setq lista-object-watched 
    (foreach par lista-pares-textos
      (progn
        (setq lista-object-watched (cons (vlax-ename->vla-object (car par)) lista-object-watched))
        (setq lista-object-watched (cons (vlax-ename->vla-object (cdr par)) lista-object-watched))
      )
    )
  )
  lista-object-watched
)


; BUILD REACTOR TEXTO MODIFICADO
; Recibe una lista de objetos VLAX que observará, y una lista de pares de textos para agregarle como data al reactor.
(defun build-reactor-texto-modificado (lista-object-watched lista-pares-textos)
  (vlr-object-reactor 
    ; OWNERS:
    lista-object-watched
    ; DATA:
    ; Data a agregada al reactor. Se adopta una lista de asociacion que tiene el nombre del reactor y la lista de pares de texto.
    (list 
      (cons "nombre-reactor" "reactor-texto-modificado")
      (cons "lista-pares-textos" lista-pares-textos)
    )
    ; EVENTO Y CALLBACK:
    '((:vlr-modified . callback-regen-db-armaduras))
  )
)


; CALLBACK PARA EL REACTOR AL EVENTO MODIFICAR 
; Callback para el reactor-texto-modificado
(defun callback-regen-db-armaduras (notifier-object object-reactor parameter-list / data lista-pares-textos par ename-texto-1 
                                    ename-texto-2 ename-notifier-object)
  ; Se lee la información almacenada en el reactor.
  (setq lista-pares-textos (cdr (assoc "lista-pares-textos" (vlr-data object-reactor))))
  ; Se ponen en cero las longitudes acumuladas para cada diametro
  (dic-armaduras-reset)
  (foreach par lista-pares-textos
    (setq ename-texto-1 (car par) ename-texto-2 (cdr par))
    ; Se agrega el par diametro-longitud a la base de datos
    (dic-armaduras-agregar ename-texto-1 ename-texto-2)
  )
  ; Se actualiza la planilla de Excel
  (dic-armaduras->excel "dic-armaduras")
  (prompt "\nSe reaccionó al evento modificar objecto.")
)


; ELIMINAR REACTORES DE NOMBRE "REACTOR-TEXTO-MODIFICADO"
(defun eliminar-reactores-texto-modificado ( / lista-object-reactors)
  ; Se obtiene una lista de reactores de tipo :vlr-object-reactor que haya en el dibujo.
  (setq lista-object-reactors (cdr (assoc :vlr-object-reactor (vlr-reactors))))
  ;  Se eliminan los viejos reactores con el nombre "reactor-texto-modificado"
  (foreach 
    reactor
    lista-object-reactors 
    (if (equal (cdr (assoc "nombre-reactor" (vlr-data reactor))) "reactor-texto-modificado")
      (vlr-remove reactor)
    )
  )
)

;#endregion


;;-----------------------------------------------------------------------------------------------------------------------------;;
;; COMANDOS PRINCIPALES DE LA APLICACION
;;-----------------------------------------------------------------------------------------------------------------------------;;
;#region COMANDOS PRINCIPALES DE LA APLICACION

(defun C:CYPE:COMPUTAR-TABLAS-ACERO ( / sset lista-entidades ename-texto-1 ename-texto-2 sub-lista-1 sub-lista-2 data-pares lista-object-watched)
  ; Se solicita al usuario seleccionar las tablas a computar.
  (alert "Seleccione tablas a computar")
  (setq sset (ssget '((0 . "TEXT"))))
  ; Se convierte el sset en una lista de enames
  (setq lista-entidades (sset-to-list sset))
  ; Se obtiene una sub-lista de los nombres de entidades que son textos y comienzan con "%%c" (diametro)
  (setq lista-entidades-diam (vl-remove-if-not 'cumple-comienza-con-fi lista-entidades))
  ; Se ponen en cero las longitudes acumuladas para cada diametro
  (dic-armaduras-reset)
  ; Se inicializa la lista de pares de texos asociados (diam . longitud)
  (setq *lista-pares-textos* nil)
  ; Para cada texto con el diametro:
  (foreach ename-texto-1 lista-entidades-diam
    (progn
      ; Se filtran todos los textos de la lista de entidades que esten alineados horizontalmente con ename-texto-1
      (setq sub-lista-1 (filtrar-alineados-hor ename-texto-1 lista-entidades))
      ; Si la sub-lista no esta vacia
      (if sub-lista-1
        ; Ahora se filtran todos los textos cuya coordenada x esté a una distancia especifica del texto de referencia ename-texto-1
        (progn
          (setq sub-lista-2 (filtrar-textos-proximos ename-texto-1 sub-lista-1))
          ; Si se encontro el par asociado a ename-texto-1
          (if sub-lista-2
            (progn
              (setq ename-texto-2 (car sub-lista-2))
              ; Se agrega el par diametro-longitud a la base de datos
              (dic-armaduras-agregar ename-texto-1 ename-texto-2)
              ; Se remueve de la lista de entidades a ename-texto-1 y ename-texto-2
              (setq lista-entidades (vl-remove ename-texto-1 lista-entidades))
              (setq lista-entidades (vl-remove ename-texto-2 lista-entidades))
              ; Se cambia el color del par de textos
              (cambiar-color-entidad ename-texto-1 4)
              (cambiar-color-entidad ename-texto-2 4)
              ; Se agrega el par a la lista de pares de textos asociados (diam . longitud)
              (setq *lista-pares-textos* (cons (cons ename-texto-1 ename-texto-2) *lista-pares-textos*))
            )
          )
        )
      )
    )
  )
  ; Se crea una lista de objetos que serán observados por el reactor "reactor-texto-modificado"
  (setq lista-object-watched (crear-lista-object-watched *lista-pares-textos*))
   ; Se eliminan del dibujo todos los reactores viejos de nombre "reactor-texto-modificado"
  (eliminar-reactores-texto-modificado)
  ; Se crea el reactor y se lo agrega a cada objeto propietario, osea, cada objeto de la lista-object-watched
  (if lista-object-watched
    (setq reactor-texto-modificado (build-reactor-texto-modificado lista-object-watched *lista-pares-textos*))
  )
  (iniciar-excel)
  (dic-armaduras->excel "dic-armaduras")
)


(defun C:CYPE:CONFIGURAR-DISTANCIA-TEXTOS ( / h-texto)
  (alert "Seleccione un texto de la columna Diametro.")
  (setq h-texto (cdr (assoc 40 (entget (ssname (ssget '((0 . "TEXT"))) 0)))))
  (alert "Seleccione la distancia horizantal entre un texto de la columna de Diam. y un texto de la columna Total")
  (vlax-ldata-put "dic-configuracion" "distancia-textos" (/ (getdist) h-texto))
  (alert "Distancia configurada correctamente.")
)


(defun C:CYPE:TERMINAR-COMPUTO-TABLAS-ACERO ( )
  ; Se eliminan del dibujo todos los reactores de nombre "reactor-texto-modificado"
  (eliminar-reactores-texto-modificado)
  ; Se restablecen los colores de los textos
  (foreach par *lista-pares-textos*
    (progn
      (quitar-color-entidad (car par))
      (quitar-color-entidad (cdr par))
    )
  )
  (terminar-excel)
)


(defun C:CYPE:COMPUTAR-ACERO-LOSAS ( / sset lista-entidades sub-lista-1 cant-barras diam-barras sep-barras long-barras string)
  ; Se solicita al usuario seleccionar las tablas a computar.
  (alert "Seleccione las losas a computar")
  (setq sset (ssget '((0 . "MTEXT"))))
  ; Se convierte el sset en una lista de enames
  (setq lista-entidades (sset-to-list sset))
  ; Se ponen en cero las longitudes acumuladas para cada diametro
  (dic-losas-reset)
  ; Se filtra una sub-lista de las entidades que cumplen con el patrón "[Cantidad] %%c[Diametro] c/[Separacion]cm L=[longitud]cm"
  (setq sub-lista-1 (foreach ename lista-entidades (vl-remove-if-not 'cumple-patron-textos-losas lista-entidades)))
  (foreach ename sub-lista-1 
    (progn
      (highlight-ent ename)
      (setq string (cdr (assoc 1 (entget ename))))
      (setq cant-barras (leer-cantidad-barras-losa string))
      (setq diam-barras (leer-diametro-barras-losa string))
      (setq sep-barras (leer-separacion-barras-losa string))
      (setq long-barras (leer-longitud-barras-losa string))
      ; Test lectura de cantidad de barras:
      ; (print (strcat string " | cantidad: " cant-barras " - diametro: " diam-barras " - separacion: " sep-barras " - longitud: " long-barras))
      ; Se agrega el par diametro-longitud a la base de datos
      (dic-losas-agregar cant-barras diam-barras sep-barras long-barras)
      ; (unhighlight-ent ename)
    )
  )
  (iniciar-excel)
  (dic-armaduras->excel "dic-losas")
)
 

(defun C:CYPE:FIX-CYPETEXT ( / sset ename ent i altura resp)
  (prompt "Seleccione los textos por ajustar: ")
  (setq sset (ssget '((0 . "TEXT"))))
  ; Se inicializa la variable altura de texto
  (setq altura 0.2)
  ; Se le solicita al usuario ingresar la altura de texto que deseee.
  (setq resp (getreal "Ingrese la altura para los textos<0.2>: "))
  ; Si la respuesta del usuario no es nula se actualiza la altura del texto.
  (if resp 
    (setq altura resp)
  )
  (setq i 0)
  ; Se itera sobre cada texto seleccionado
  (while (< i (sslength sset))
    (setq ename (ssname sset i))
    (setq ent (entget ename))
    ; Se cambia el tipo de justificación del texto (cualquiera que tenga) por el tipo 
    ; de justificación "Izquierda", el código DFX y su valor es (72 . 0).
    ; Por lo general, los textos del Cype suelen tener el tipo de justificación "Ajustar" 
    ; esto corresponde al par punteao: (72 . 5)
    (setq ent (entmod (subst (cons 72 0) (assoc 72 ent) ent)))
    ; Se cambia la altura del texto, el código DFX asociado es el (40 . altura_texto)
    (setq ent (entmod (subst (cons 40 altura) (assoc 40 ent) ent)))
    ; Se establece el factor de escala X relativa. Como el texto del Cype suele
    ; tener justificación ajustado, el valor del factor de escala tiene valores muy 
    ; diferentes. Se lo establecerá en 1. El código DFX para este valor es (41 . factor_escala)
    (setq ent (entmod (subst (cons 41 1) (assoc 41 ent) ent)))
    ; Se actualiza la imagen de la entidad en la pantalla.
    (entupd ename)
    (setq i (1+ i))
  )
)


;#endregion


;;-----------------------------------------------------------------------------------------------------------------------------;;
;; CUERPO PRINCIPAL DE LA APLICACION - MAIN
;;-----------------------------------------------------------------------------------------------------------------------------;;
;#region
; Se define la distnacia por defecto entre textos de la columna "Diam" y la columna "Total (cm)"
(vlax-ldata-put "dic-configuracion" "distancia-textos" 30.165)

;#endregion

