
(defun C:CYPE:COMPUTAR-ACERO()
  (alert "Seleccione tablas a computar")
  (setq sset (ssget '((0 . "TEXT"))))
  (sslength sset)
)


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