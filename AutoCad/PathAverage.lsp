;;  PathAverage.lsp [command name: PAv]
;;  To make a Path that is the Average between two selected path-type objects.
;;  Asks User for number of increments to divide paths into, and draws a
;;    LWPolyline along the halfway points between those subdivision points.
;;  Paths may be of different entity types, open or closed, and if both Polylines,
;;    may be of different varieties and/or have different numbers of vertices.
;;  Draws on current Layer, with current PLINEWID System Variable setting,
;;    in all line segments [PEDIT for curvature if desired].
;;  Remembers number of increments and offers as default on subsequent use.
;;  Kent Cooper, 21 October 2014

;;;;; Currently doesn't account for 3rd dimension -- future upgrade: make 3D
;;;;;   Polyline if selected paths are not in same plane or parallel planes.
;;;;; Could force PLINEWID = 0, or same as selected paths if both are Polylines
;;;;;   with same width.
;;;;; Could draw on Layer of paths if they are on the same Layer, or ask User
;;;;;   for Layer to draw on.
;;;;; Does not account for paths running in opposite directions -- if needed, see
;;;;;   http://cadtips.cadalyst.com/object-properties/reverse-entity-direction
;;;;;   to reverse one, including more entity types than AutoCAD's Reverse
;;;;;   command will do.

(defun C:PAv ; = Path Average
  (/ *error* var ev getobj avgpt path1 path2 seg1 seg2 clo1 clo2 cmde osm inc)
  (defun *error* (errmsg)
    (if (not (wcmatch errmsg "Function cancelled,quit / exit abort,console break"))
      (princ (strcat "\nError: " errmsg))
    ); if
    (setvar 'cmdecho cmde)
    (setvar 'osmode osm)
    (princ)
  ); defun - *error*

  (defun var (typ); build variable name with number
    (read (strcat typ num))
  ); defun - var

  (defun ev (typ); evaluate what's in variable name with number
    (eval (read (strcat typ num)))
  ); defun - ev

  (defun getobj (num / esel edata etype)
    (while
      (not
        (and
          (setq
            esel (entsel (strcat "\nSelect Path object #" num " for Averaging: "))
            edata (if esel (entget (car esel)))
            etype (if esel (cdr (assoc 0 edata)))
          ); setq
          (wcmatch etype "LINE,*POLYLINE,ARC,CIRCLE,ELLIPSE,SPLINE")
        ); and
      ); not
      (prompt "\nNothing selected, or not a path-like object --")
    ); while
    (set (var "path") (car esel)); path1 or path2
    (set (var "seg"); seg1 or seg2 [length of subdividing segment of path]
      (/ (vlax-curve-getDistAtParam (ev "path") (vlax-curve-getEndParam (ev "path"))) *PAvsegs)
    ); set
    (set (var "clo") (vlax-curve-isClosed (ev "path"))); clo1 or clo2
  ); defun -- getobj

  (defun avgpt (n)
    (mapcar '/
      (mapcar '+
        (vlax-curve-getPointAtDist path1 (* seg1 n))
        (vlax-curve-getPointAtDist path2 (* seg2 n))
      ); mapcar +
      '(2 2 2)
    ); mapcar /
  ); defun -- avgpt

  (setq
    cmde (getvar 'cmdecho)
    osm (getvar 'osmode)
    inc -1
  ); setq
  (setvar 'cmdecho 0)
  (setvar 'osmode 0)
  (initget (if *PAvsegs 6 7)); no zero, no negative; no Enter on first use
  (setq *PAvsegs ; global variable for subsequent default
    (cond
      ( ; User input
        (getint
          (strcat
            "\nNumber of Segments to divide Paths into"
            (if *PAvsegs (strcat " <" (itoa *PAvsegs) ">: ") ": ")
          ); strcat
        ); getint
      ); User input [other than Enter] condition
      (*PAvsegs); Enter on subsequent use
    ); cond
  ); setq

  (getobj "1")
  (redraw path1 3); highlight
  (getobj "2")
  (redraw path1 4); unhighlight
  (command "_.pline")
  (repeat *PAvsegs
    (command (avgpt (setq inc (1+ inc)))); start of each segment
  ); repeat
  (if (and clo1 clo2)
    (command "_close")
    (command
      (mapcar '/
        (mapcar '+ (vlax-curve-getEndPoint path1) (vlax-curve-getEndPoint path2))
        '(2 2 2)
      ); mapcar /
      ""
    ); command
  ); if
  (setvar 'cmdecho cmde)
  (setvar 'osmode osm)
  (princ)
); defun -- C:PAv

(vl-load-com)
(prompt "\nType PAv to draw an Average Path between two selected Paths.")