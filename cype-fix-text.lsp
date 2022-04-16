
(defun C:FIX-CYPETEXT ( / sset ename ent i altura resp)
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
