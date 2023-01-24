(defun c:max-inscribed-rect-each ( / *error* unit unique GroupPoints ClosestInters ClosestInters+line LWPoly midpts segcalc ss i obj acadmodel osm maxRect seg pl arrecs)

(vl-load-com)

(setq *variables* '(unit unique GroupPoints ClosestInters ClosestInters+line LWPoly midpts segcalc ss i obj acadmodel osm maxRect seg pl arrecs))

(defun *error* (msg)
(cond
((not (vl-position msg '("Function cancelled" "*Cancel*" "quit / exit abort")))
(princ "\nAn error has occured.")
)
(T (princ "\nExit."))
)

(if osm (setvar 'osmode osm))
(if (/= *variables* nil) (mapcar '(lambda (x) (set x nil)) *variables*))
(setq *variables* nil)
(vla-endundomark (vla-get-activeDocument (vlax-get-acad-object)))
(redraw)

(gc)
(princ)
);defun

(defun unit (v / d)
(if (not (equal (setq d (distance '(0.0 0.0 0.0) v)) 0.0 1e-8))
(mapcar '(lambda (x) (/ x d)) v)
)
);defun ;; mod by M.R.

(defun unique (lst)
(if lst (cons (car lst) (unique (vl-remove-if '(lambda (x) (equal x (car lst) 1e-6)) lst))))
);defun ;; mod by M.R.

(defun GroupPoints (lst)
(if (> (length lst) 2)
(cons (list (car lst) (cadr lst) (caddr lst)) (groupPoints (cdddr lst)))
)
);defun

(defun ClosestInters (ray pline pt)
(cadr (vl-sort (GroupPoints (vlax-invoke ray 'intersectWith pline acExtendNone))
'(lambda (a b) (< (distance pt a) (distance pt b)))))
);defun

(defun ClosestInters+line (ray pline line pt)
(cadr (vl-sort (unique (GroupPoints (append (vlax-invoke ray 'intersectWith pline acExtendNone) (vlax-invoke ray 'intersectWith line acExtendNone))))
'(lambda (a b) (< (distance pt a) (distance pt b)))))
);defun ;; mod by M.R.

(defun LWPoly (lst)
(entmakex (append
(list (cons 0 "LWPOLYLINE")
(cons 100 "AcDbEntity")
(cons 100 "AcDbPolyline")
(cons 90 (length lst))
(cons 70 (1+ (* (getvar 'plinegen) 128)))
)
(mapcar (function (lambda (p) (cons 10 p))) lst)
(list (list 210 0.0 0.0 1.0))
)
)
);defun

(defun midpts (obj / k p pl)
(setq k -0.5)
(repeat (fix (vlax-curve-getEndParam obj))
(setq p (vlax-curve-getPointAtParam obj (setq k (1+ k))))
(setq pl (cons p pl))
)
(reverse pl)
)

(defun segcalc (obj segPt seg / segPtPar segStartPt segEndPt segAngle dir vl ray inter li step stepBase stepForward stepBack ray2 bp1 bp2 p1 p2 len width d maxArea maxRect)
(setq segPtPar (vlax-curve-getParamAtPoint obj (vlax-curve-getClosestPointTo obj segPt)))
(setq segStartPt (vlax-curve-getPointAtParam obj (fix segPtPar)))
(setq segEndPt (vlax-curve-getPointAtParam obj (+ (fix segPtPar) 1)))
(setq segAngle (angle segStartPt segEndPt))
(setq dir 1)
(setq vl (mapcar '(lambda (p) (trans (list (car p) (cadr p) 0.0) (vlax-vla-object->ename obj) 0)) (mapcar 'cdr (vl-remove-if '(lambda (x) (/= (car x) 10)) (entget (vlax-vla-object->ename obj)))))) ;; mod by M.R.

(setq ray (vla-addray acadmodel
		      (vlax-3D-point (vlax-curve-getPointAtParam obj (+ (fix segPtPar) 0.5))) ;; mod by M.R.
		      (vlax-3D-point (polar (vlax-curve-getPointAtParam obj (+ (fix segPtPar) 0.5)) (+ segAngle (/ pi 2.0 dir)) 10))))
(setq inter (vlax-invoke ray 'intersectWith obj acExtendNone))
(if (> (length inter) 3) ;; mod by M.R.
(setq dir 1)
(setq dir -1)
)
(vla-delete ray)

(setq ray (vla-addray acadmodel (vlax-3d-point segStartPt) (vlax-3d-point (polar segStartPt (+ segAngle pi) 10)))) ;; mod by M.R.
(setq segStartPt (if (ClosestInters ray obj segStartPt) (ClosestInters ray obj segStartPt) segStartPt)) ;; mod by M.R.
(vla-delete ray) ;; mod by M.R.

(setq ray (vla-addray acadmodel (vlax-3d-point segEndPt) (vlax-3d-point (polar segEndPt segAngle 10)))) ;; mod by M.R.
(setq segEndPt (if (ClosestInters ray obj segEndPt) (ClosestInters ray obj segEndPt) segEndPt)) ;; mod by M.R.
(vla-delete ray) ;; mod by M.R.

(setq li (vla-addline acadmodel (vlax-3d-point segStartPt) (vlax-3d-point segEndPt))) ;; mod by M.R.

(setq step (/ (distance segStartPt segEndPt) seg))
(setq stepBase (vlax-3D-point '(0 0 0)))
(setq stepForward (vlax-3D-point (polar '(0 0 0) segAngle step)))
(setq stepBack (vlax-3D-point (polar '(0 0 0) (+ segAngle pi) step)))

(setq ray (vla-addray acadmodel (vlax-3D-point segStartPt) (vlax-3D-point (polar segStartPt (+ segAngle (/ pi 2.0 dir)) 10))))
(while (equal segStartPt (vlax-invoke ray 'intersectWith obj acExtendNone) 1e-6) ;; mod by M.R.
(vla-move ray stepBase (vlax-3D-point (polar '(0 0 0) segAngle step))) ;; mod by M.R.
)

(setq ray2 (vla-addray acadmodel (vlax-3D-point segEndPt) (vlax-3D-point (polar segEndPt (+ segAngle (/ pi 2.0 dir)) 10))))
(while (equal segEndPt (vlax-invoke ray2 'intersectWith obj acExtendNone) 1e-6) ;; mod by M.R.
(vla-move ray2 stepBase (vlax-3D-point (polar '(0 0 0) (+ segAngle pi) step))) ;; mod by M.R.
)

(setq bp1 (vlax-get ray 'basePoint))
(setq bp2 (vlax-get ray2 'basePoint))
(setq p1 bp1)
(setq p2 bp2)
(setq maxArea 0.0)

(while (and (vlax-curve-getParamAtPoint li p1) (vlax-curve-getParamAtPoint li p2) (> (vlax-curve-getParamAtPoint li p2) (vlax-curve-getParamAtPoint li p1))) ;; mod by M.R.
(while (and (vlax-curve-getParamAtPoint li p1) (vlax-curve-getParamAtPoint li p2) (> (vlax-curve-getParamAtPoint li p2) (vlax-curve-getParamAtPoint li p1))) ;; mod by M.R.
(setq len (distance p1 p2))
(setq width (min (distance (ClosestInters+line ray obj li p1) p1) (distance (ClosestInters+line ray2 obj li p2) p2))) ;; mod by M.R.
(foreach v vl
(if (equal (setq d (distance p1 (vlax-curve-getclosestpointto ray v))) 0.0 1e-6)
(setq vl (vl-remove v vl))
(if (and (equal (unit (mapcar '- v (vlax-curve-getclosestpointto ray v))) (mapcar '- (unit (mapcar '- v (vlax-curve-getclosestpointto ray2 v)))) 1e-6) (< d width))
(setq width d)
)
)
);foreach ;; mod by M.R.
(if (> (* len width) maxArea)
(setq maxArea (* len width)
      maxRect (list p1 p2 (polar p2 (+ segAngle (/ pi 2.0 dir)) width) (polar p1 (+ segAngle (/ pi 2.0 dir)) width))
      )
)
(vla-move ray2 stepBase stepBack)
(setq p2 (polar p2 (+ segAngle pi) step))
);while
(vla-move ray stepBase stepForward)
(setq p1 (polar p1 segAngle step))
(vla-move ray2 (vla-get-basePoint ray2) (vlax-3D-point bp2))
(setq p2 bp2)
);while

(vla-delete ray)
(vla-delete ray2)
(vla-delete li)
maxRect
);defun

(vla-startUndoMark (vla-get-activeDocument (vlax-get-acad-object)))
(setq acadmodel (vla-get-modelspace (vla-get-activeDocument (vlax-get-acad-object))))
(setq osm (getvar 'osmode))
(setvar 'osmode 0)
(initget 6)
(setq seg (getint "\nSpecify segmentation per segment - must be > 1... <50> : "))
(if (null seg)
(setq seg 50)
)
(while (= seg 1)
(initget 6)
(setq seg (getint "\nSpecify segmentation per segment - must be > 1... <50> : "))
(if (null seg)
(setq seg 50)
)
)
(if (setq ss (ssget '((0 . "LWPOLYLINE") (-4 . "&=") (70 . 1) (-4 . "<not") (-4 . "<>") (42 . 0.0) (-4 . "not>"))))
(repeat (setq i (sslength ss))
(setq obj (vlax-ename->vla-object (ssname ss (setq i (1- i)))))
(setq pl (midpts obj))
(foreach p pl
(setq maxRect (segcalc obj p (if (<= seg 50) seg 50)))
(setq arrecs (cons (list (* (distance (car maxRect) (cadr maxRect)) (distance (cadr maxRect) (caddr maxRect))) maxRect p) arrecs))
)
(setq maxRect (car (vl-sort arrecs '(lambda (a b) (> (car a) (car b))))))
(if (<= seg 50)
(progn
(LWPoly (cadr maxRect))
(vla-put-color (vlax-ename->vla-object (entlast)) 6)
(setq maxRect nil arrecs nil)
)
(progn
(setq maxRect (segcalc obj (caddr maxRect) seg))
(LWPoly maxRect)
(vla-put-color (vlax-ename->vla-object (entlast)) 6)
(setq maxRect nil arrecs nil)
)
)
)
)
(setvar 'osmode osm)
(vla-endUndoMark (vla-get-activeDocument (vlax-get-acad-object)))
(setq *variables* nil)
(princ)
);defun