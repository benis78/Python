;***SCRIPTABLE ETRANSMIT

(apply

'(lambda ()

(SETVAR "CMDECHO" 0)

(COMMAND "QSAVE" )

;;;*-----------------------------------------------------------

(COMMAND "LAYER" "UNLOCK" "*" "")

(COMMAND "TILEMODE" "0" )

(COMMAND "pspace" )

(command "LAYER" "t" "VIEWPORT" "u" "*" "m" "VIEWPORT" "c" "8" "VIEWPORT" "T" "0" "S" "0" "")

(if(="Model" (getvar "ctab"))

(setq XX1 (ssget "_X" '((0 . "viewport"))))

(progn

(foreach mb (vl-remove "Model"(layoutlist))

(setvar "ctab" mb)

(setvar "psltscale" 1)

(setvar "ltscale" 1)

(setq XX1 (ssget "_X" '((0 . "viewport"))))

(command "-vports" "lock" "off" xx1 ""))))

;;;*-----------------------------------------------------------

(defun removexref (xrefname / blkname)

(if (setq blkname (tblsearch "block" xrefname))

(if (= (cdr (assoc 70 blkname)) 12)

(command ".xref" "d" xrefname)

)

)

(princ)

)

(defun remove-unloaded-xrefs ()

(vlax-for block (vla-get-blocks

(vla-get-activedocument

(vlax-get-acad-object)))

(if (and (= :vlax-true (vla-get-isxref block))

(= 0 (vla-get-count block))

)

(vla-detach block)

)

)

)

(apply

'(lambda ()

(remove-unloaded-xrefs)

(princ)

)

'()

)

;;;*-----------------------------------------------------------

(defun TODAY ( / d yr mo da)

(setq d (rtos (getvar "CDATE") 2 6)

yr (substr d 3 2)

mo (substr d 5 2)

DAY (substr d 7 2)

);setq

(strcat yr mo DAY)

);defun

;;;*-----------------------------------------------------------

(defun NOW ( / d hr mi se)

(setq d (rtos (getvar "CDATE") 2 6)

hr (substr d 10 2)

mi (substr d 12 2)

se (substr d 14 2)

);setq

(strcat hr mi se)

);defun

;;;*-----------------------------------------------------------

(SETQ CPFIX (getvar "dwgprefix"))

(SETQ CDNAM (getvar "dwgNAME"))

(SETQ NDIR (STRCAT "eTran-" (today) "/" ))

(vl-mkDir (strcat CPFIX NDIR ))

;;;*-----------------------------------------------------------

(SETQ NFNAME (strcat (vl-filename-base (getvar "dwgname")) "-" (today) (NOW) ))

;;;*------------------------------------------------------------

(command "saveas" "2004" (strcat (getvar "dwgprefix") NDIR NFNAME ".dwg"))

;;;*------------------------------------------------------------

(command "ETRANSMIT" "CH" "ETRANS" "C" (strcat (getvar "dwgprefix") NFNAME ))

;;;*------------------------------------------------------------

(COMMAND "QSAVE" )

;;;*------------------------------------------------------------

(princ)

)

'()

)