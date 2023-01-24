(defun c:test ( / *error* GroupPoints ClosestInters LWPoly ent obj acadmodel segPt segPtPar segStartPt segEndPt segAngle
	       dir step stepBase stepForward stepBack osm ray inter ray2 bp1 bp2 p1 p2 len width maxArea maxRect)

(setq *variables* '(GroupPoints ClosestInters LWPoly ent obj acadmodel segPt segPtPar segStartPt segEndPt segAngle
		    dir step stepBase stepForward stepBack osm ray inter ray2 bp1 bp2 p1 p2 len width maxArea maxRect))

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

(defun GroupPoints (lst)
(if (> (length lst) 2)
(cons (list (car lst) (cadr lst) (caddr lst)) (groupPoints (cdddr lst)))
)
);defun

(defun ClosestInters (ray pline pt)
(cadr (vl-sort (GroupPoints (vlax-invoke ray 'intersectWith pline acExtendNone))
'(lambda (a b) (< (distance pt a) (distance pt b)))))
);defun

(defun LWPoly (lst)
(entmakex (append (list (cons 0 "LWPOLYLINE")
(cons 100 "AcDbEntity")
(cons 100 "AcDbPolyline")
(cons 90 (length lst))
(cons 70 1))
(mapcar (function (lambda (p) (cons 10 p))) lst)))
);defun


(while (not ent) (setq ent (entsel "\nSelect a closed polyline: ")))
(setq obj (vlax-ename->vla-object (car ent)))

(if (and (= (vla-get-objectName obj) "AcDbPolyline")
	 (= (vla-get-closed obj) :vlax-true)
	 )
(progn
(vla-startUndoMark (vla-get-activeDocument (vlax-get-acad-object)))
(setq acadmodel (vla-get-modelspace (vla-get-activeDocument (vlax-get-acad-object))))

(setq segPt (getpoint "\nPick a point, near one of the PLINE segments."))
(setq segPtPar (vlax-curve-getParamAtPoint obj (vlax-curve-getClosestPointTo obj segPt)))
(setq segStartPt (vlax-curve-getPointAtParam obj (fix segPtPar)))
(setq segEndPt (vlax-curve-getPointAtParam obj (+ (fix segPtPar) 1)))
(setq segAngle (angle segStartPt segEndPt))
(setq dir 1)
(setq step (/ (distance segStartPt segEndPt) 50))
(setq stepBase (vlax-3D-point '(0 0 0)))
(setq stepForward (vlax-3D-point (polar '(0 0 0) segAngle step)))
(setq stepBack (vlax-3D-point (polar '(0 0 0) (+ segAngle pi) step)))
(setq osm (getvar 'osmode))
(setvar 'osmode 0)

(setq ray (vla-addray acadmodel
		      (vlax-3D-point (polar (vlax-curve-getPointAtParam obj (+ (fix segPtPar) 0.5)) (+ segAngle (/ pi 2.0 dir)) 0.1))
		      (vlax-3D-point (polar (vlax-curve-getPointAtParam obj (+ (fix segPtPar) 0.5)) (+ segAngle (/ pi 2.0 dir)) 10))))
(setq inter (vlax-invoke ray 'intersectWith obj acExtendNone))
(if (and (> (length inter) 0) (/= (rem (/ (length inter) 3) 2) 0))
(setq dir 1)
(setq dir -1)
)
(vla-delete ray)

(setq ray (vla-addray acadmodel (vlax-3D-point segStartPt) (vlax-3D-point (polar segStartPt (+ segAngle (/ pi 2.0 dir)) 10))))
(while (/= (rem (/ (length (vlax-invoke ray 'intersectWith obj acExtendNone)) 3) 2) 0)
(vla-move ray stepBase (vlax-3D-point (polar '(0 0 0) segAngle 0.1)))
)

(setq ray2 (vla-addray acadmodel (vlax-3D-point segEndPt) (vlax-3D-point (polar segEndPt (+ segAngle (/ pi 2.0 dir)) 10))))
(while (/= (rem (/ (length (vlax-invoke ray2 'intersectWith obj acExtendNone)) 3) 2) 0)
(vla-move ray2 stepBase (vlax-3D-point (polar '(0 0 0) (+ segAngle pi) 0.1)))
)

(setq bp1 (vlax-get ray 'basePoint))
(setq bp2 (vlax-get ray2 'basePoint))
(setq p1 bp1)
(setq p2 bp2)

(while (and (vlax-curve-getParamAtPoint obj p1) (> (vlax-curve-getParamAtPoint obj p2) (vlax-curve-getParamAtPoint obj p1)))

(while (> (vlax-curve-getParamAtPoint obj p2) (vlax-curve-getParamAtPoint obj p1))

(setq len (distance p1 p2))
(setq width (min (distance (ClosestInters ray obj p1) p1) (distance (ClosestInters ray2 obj p2) p2)))
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
(LWPoly maxRect)
(vla-put-color (vlax-ename->vla-object (entlast)) 6)

(setvar 'osmode osm)
(vla-endUndoMark (vla-get-activeDocument (vlax-get-acad-object)))
)
(princ "\nThe selected object is not a closed polyline.")
);if closed pl

(setq *variables* nil)
(princ)
);defun