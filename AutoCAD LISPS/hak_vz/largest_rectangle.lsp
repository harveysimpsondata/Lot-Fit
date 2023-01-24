
(defun c:maxalignedrect
( /
    *error* getintersections mappend mklist flatten revlwpline to_counter_clockwise pointlist2d pointlist3d vlax-make-array-type
    sort_nearest convert_to_vl_object perperdicular_from_point_to_line unique doc modelspace e osm poly_obj poly_cords seglist i p1 p2 p3 p4 ang_ortho ang_p1_p2 ang_p2_p1 dist
    max_rect max_rect_area line1 line2 int_1 int_2 dist dist_min total_dist step ar _area edges edge_seg test_line
)
; maxalignedrect - creates rectangle with "approximately" largest area aligned to one edge of enclosing polygon
;
; Author: Miljenko Hatlak (hak_vz)
; https://forums.autodesk.com/t5/user/viewprofilepage/user-id/5530556
; 27.03.2020. 
; Code is working correctly for polygons created as lwpolyline objects with line and arc segments.
; Reshape your polygons so that there is no unnecessary vertices inside straight line segments or arcs.
; If you have to align rectangle to an arc you should temporary decurve enclosing polygon.
; Final rectangle may be created so that it intersect with enclosing polygon - testing is not
; implemented since it would hugely affect speed of execution. 
; Avoid using it with oddly shaped polygons.
; 

(defun *error* (msg)
  (setvar "cmdecho" 1)
  (if (not (wcmatch (strcase msg) "*BREAK*,*CANCEL*,*EXIT*"))
    (progn
       (setvar 'cmdecho 1)
       (setvar 'osmode osm)
       (vla-endundomark doc)
       (princ (strcat "\nOops an Error : ( " msg " ) occurred."))
    )
  ) 
  (princ)
 )
 
(defun getintersections	(obj1 obj2 / var)
     ; from RonJonP
     (setq var (vlax-variant-value (vla-intersectwith obj1 obj2 0)))
     (if (< 0 (vlax-safearray-get-u-bound var 1))
	  (vlax-safearray->list var)
     ) ;_ end of if
) ;_ end of defun
(defun mappend (fn lst) ; Peter Norvig ??
     ; Append the results of calling fn on each element of list.
     ; Like mapcon, but uses append instead of nconc."
     ; One thing to notice is that fn must return a list, otherwise, it will go wrong.
     ; usage: (mappend '(lambda (x) (list x (* x x))) '(1 2 3))
     (apply 'append (mapcar fn lst))
) ;_ end of defun
(defun mklist (x)
     ; If x is a list return it, otherwise return the list of x
     (if (listp x)
	  x
	  (list x)
     ) ;_ end of if
) ;_ end of defun

(defun flatten (expr)
     ; Get rid of imbedded lists (to one level only)."
     (mappend 'mklist expr)
) ;_ end of defun

(defun revlwpline (e / footer done vertices header flag)
  ;reverse lightweight polyline
  ;http://www.metzgerwillard.us/tdavis/lisp/reverse.html
  (foreach item (reverse (entget e))
    (cond
      ((not done)
        (cond
          ((= (car item) 40)
            (setq footer (cons (cons 41 (cdr item)) footer)      ;swap width
                  done t
            )
          )
          ((= (car item) 41)
            (setq footer (cons (cons 40 (cdr item)) footer))     ;swap width
          )
          ((= (car item) 42)
            (setq footer (cons (cons 42 (- (cdr item))) footer)) ;negate bulge
          )
          ((= (car item) 210)
            (setq footer (cons item footer))
          )
        )
      )
      ((= (car item) 10)
        (setq vertices (cons item vertices))
      )
      ((= (car item) 40)
        (setq vertices (cons (cons 41 (cdr item)) vertices))     ;swap width
      )
      ((= (car item) 41)
        (setq vertices (cons (cons 40 (cdr item)) vertices))     ;swap width
      )
      ((= (car item) 42)
        (setq vertices (cons (cons 42 (- (cdr item))) vertices)) ;negate bulge
      )
      (t (setq header (cons item header)))
    )
  )
  (setq flag (assoc 70 header))
  (if (< (cdr flag) 128)                 ;turn on linetype generation
    (setq header (subst (cons 70 (+ (cdr flag) 128)) flag header))
  )
  (entmod (append header (reverse vertices) footer))
  (princ)
)

(defun to_counter_clockwise ( e / LW LST MAXP MINP)
; Writer Evgeniy Elpanov.
; modified ny hak_vz
  (setq lw (vlax-ename->vla-object e))
  (vla-GetBoundingBox lw 'MinP 'MaxP)
  (setq
      minp (vlax-safearray->list minp)
      MaxP (vlax-safearray->list MaxP)
      lst 
        (mapcar
          (function
          (lambda (x)
          (vlax-curve-getParamAtPoint
          lw
          (vlax-curve-getClosestPointTo lw x)
          ) ;_ vlax-curve-getParamAtPoint
          ) ;_ lambda
          ) ;_ function
          (list minp
               (list (car minp) (cadr MaxP))
                MaxP
               (list (car MaxP) (cadr minp))
                ) ;_ list
          ) ;_ mapcar
      ) ;_ setq
  (if 
    (or
      (<= (car lst) (cadr lst) (caddr lst) (cadddr lst))
      (<= (cadr lst) (caddr lst) (cadddr lst) (car lst))
      (<= (caddr lst) (cadddr lst) (car lst) (cadr lst))
      (<= (cadddr lst) (car lst) (cadr lst) (caddr lst))
     ) ;_ or
        (revlwpline e)
    ) ;_ if

) ;_ defun

(defun pointlist2d (lst / ret)
 (while lst(setq	ret (cons (list (car lst) (cadr lst)) ret) lst (cddr lst))) 
 (reverse ret)
)
(defun pointlist3d (lst / ret)
     ; converts one dimensional list (vector) to list of 3d points 
     (while lst
	  (setq	ret (cons (list (car lst) (cadr lst) (caddr lst)) ret)
		lst (cdddr lst)
	  ) ;_ end of setq
     ) ;_ end of while
     (reverse ret)
) ;_ end of defun

(defun vlax-make-array-type  (lst atype)
  (vlax-safearray-fill
    (vlax-make-safearray
      atype
      (cons 0 (1- (length lst))))
    lst))

(defun sort_nearest (pt ptlist)
     ; sort point list according to distance to point pt
     (car
	  (vl-sort ptlist
		   (function (lambda (a b) (< (distance pt a) (distance pt b)))
		   ) ;_ end of function
	  ) ;_ end of vl-sort
     ) ;_ end of car
) ;_ end of defun

(defun convert_to_vl_object (e) (vlax-ename->vla-object e))

(defun perperdicular_from_point_to_line (lin1 lin2 p / x1 y1 x2 y2 x3 y3 k m n ret)
;returns point on a line (line1 line2) as a perpendicular projection from point p
	(mapcar 'set '(x1 x2 x3) (mapcar 'car (list lin1 lin2 p)))
	(mapcar 'set '(y1 y2 y3) (mapcar 'cadr (list lin1 lin2 p)))
	(setq 
		m (-(*(- y2 y1) (- x3 x1))(*(- x2 x1) (- y3 y1)))
		n (+(* (- y2 y1)(- y2 y1))(*(- x2 x1)(- x2 x1)))
	)
	(cond 
		((/= n 0.0) 
			(setq 
				k (/ m n)
				ret (list(- x3 (* k(- y2 y1)))(+ y3 (* k(- x2 x1))))
			)
		)
	)
	ret
)

(defun unique ( l ) (if l (cons (car l) (unique (vl-remove-if '(lambda ( x ) (equal (car l) x 1e-6)) l)))))
; (defun projections_to_line_segment (line1 line2 pointlist)

; )







(vl-load-com) 
(setq doc (vla-get-activedocument (vlax-get-acad-object)))
(setq modelspace (vla-get-modelspace doc))
(vla-endundomark doc) 
(vla-startundomark doc) 
(setq e (car(entsel "\nSelect poligon >")))
(to_counter_clockwise e)
(setq osm (getvar 'osmode))
(setq poly_obj(convert_to_vl_object e))
(setvar 'osmode 512)
(setq p3 (getpoint "\nSelect a point on a polygon side that rectangle edge is aligned to > "))
(setvar 'osmode 0)
(setvar 'cmdecho 0)
(setq seglist nil i 0)
(setq poly_cords (pointlist2d(vlax-get poly_obj 'coordinates)))
(setq poly_cords (append poly_cords (list (car poly_cords))))
(while (< i (- (length poly_cords) 1)) (setq seglist(cons (list (nth i poly_cords) (nth (+ i 1) poly_cords)) seglist) i (+ i 1))) 
(foreach seg seglist (setq p1 (car seg) p2 (cadr seg))(if (equal (+ (distance p1 p3) (distance p2 p3))(distance p1 p2) 1e-3 )(setq edge_seg seg)))
(setq p1 (car edge_seg) p2 (cadr edge_seg))
(setq ang_p1_p2 (angle p1 p2))
(setq ang_ortho (+ (angle p1 p2) (/ pi 2)))
(setq ang_p2_p1 (angle p2 p1))
(setq dist (distance p1 p2) step (/ dist 50))
(setq max_rect nil)
(setq max_rect_area 0)
(setq line1 (vla-addline modelspace (vlax-3d-point (polar p1 ang_ortho 0.5) )(vlax-3d-point (polar p1 ang_ortho 1e6))))
(setq line2 (vla-addline modelspace (vlax-3d-point (polar p2 ang_ortho 0.5) )(vlax-3d-point (polar p2 ang_ortho 1e6))))

(while (not (setq int_1(getintersections line1 poly_obj)))
    (if line1(vla-delete line1))
    (setq p1 (polar p1 ang_p1_p2 step))
    (setq line1 (vla-addline modelspace (vlax-3d-point (polar p1 ang_ortho 0.5) )(vlax-3d-point (polar p1 ang_ortho 1e6))))
)
(setq int_1 (pointlist3d int_1))


(while (not (setq int_2(getintersections line2 poly_obj)))
    (if line2(vla-delete line2))
    (setq p2 (polar p2 ang_p2_p1 step))
    (setq line2 (vla-addline modelspace (vlax-3d-point (polar p2 ang_ortho 0.5) )(vlax-3d-point (polar p2 ang_ortho 1e6))))
)

(setq int_2 (pointlist3d int_2))

(setq int_1 (sort_nearest p1 int_1) int_2 (sort_nearest p2 int_2))
(if line1(vla-delete line1))
(if line2(vla-delete line2))

(setq
    total_dist (distance p1 p2)
    dist total_dist
    step (/ dist 50)
    i 0
)
(while (>= total_dist 0)

(setq line1 (vla-addline modelspace (vlax-3d-point (polar p1 ang_ortho 0.5) )(vlax-3d-point (polar p1 ang_ortho 1e6))))
(setq int_1 (sort_nearest p1 (pointlist3d (getintersections line1 poly_obj))))
(vla-delete line1)

    (while (>= dist step)
        (setq line2 (vla-addline modelspace (vlax-3d-point (polar p2 ang_ortho 0.5) )(vlax-3d-point (polar p2 ang_ortho 1e6))))
        (setq int_2 (sort_nearest p2 (pointlist3d (getintersections line2 poly_obj))))
        (vla-delete line2)
        (if (and int_1 int_2)
            (progn
            (setq dist_min (min(distance p1 int_1)(distance p2 int_2)))
            (setq p4 (polar p1 ang_ortho dist_min) p3 (polar p2 ang_ortho dist_min))
            (command "area" p1 p2 p3 p4 "")
            (setq ar (getvar 'area))
            (if (>=  ar max_rect_area)
                (progn
                    (setq test_line (vla-addline modelspace (vlax-3d-point p3)(vlax-3d-point p4)))
                    (setq int (vlax-invoke test_line 'IntersectWith poly_obj 0))
                    
                        (if (< (length int) 6) (setq max_rect_area ar max_rect (list p1 p2 p3 p4)))
                    
                    (vla-delete test_line)
                )
            )
            )
        )
         (setq p2 (polar p1 ang_p1_p2 (- dist step)))
         (setq dist (- dist step))
    )
 (setq p1 (polar p1 ang_p1_p2 step))
 (setq p2 (polar p1 ang_p1_p2 (- total_dist step)))
 (setq total_dist (- total_dist step))
 (setq dist (distance p1 p2))
 )
 
(mapcar 'set '(p1 p2 p3 p4) max_rect)
(command "pline" p1 p2 p3 p4 "c")
(setq poly_obj (convert_to_vl_object(entlast)))
(setq poly_cords (pointlist2d(vlax-get poly_obj 'coordinates)))
(setq i 0 seglist nil)
(while (< i (- (length poly_cords) 1)) (setq seglist(cons (list (nth i poly_cords) (nth (+ i 1) poly_cords)) seglist) i (+ i 1))) 
(foreach seg seglist (setq p1 (car seg) p2 (cadr seg))(if (equal (+ (distance p1 p3) (distance p2 p3))(distance p1 p2) 1e-3 )(setq edge_seg seg)))
(while (< i (- (length poly_cords) 1)) (setq seglist(cons (list (nth i poly_cords) (nth (+ i 1) poly_cords)) seglist) i (+ i 1))) 
(setq _area (vlax-get poly_obj 'area))
(setq edges (mapcar '(lambda (x) (distance (car x) (cadr x))) seglist))
(setq edges (vl-sort  edges (function (lambda (a b) (< a b)))))
(alert (strcat "Max rectangle sizes is aprox " (rtos (car edges) 2 2) " x " (rtos (last edges) 2 2) " units and area = " (rtos _area 2 2) " units sq. \n Recheck and correct if needed !"))
(setvar 'cmdecho 1)
(setvar 'osmode osm)
(vla-endundomark doc) 
(princ)
)
(princ "\nCommand maxalignerect - to draw max area rectangle insade a polygon")
(princ)

(defun list->binoms (lst / ret) (while lst (setq ret (cons (list (car lst) (cadr lst)) ret) lst (cddr lst))) (reverse ret)) 
