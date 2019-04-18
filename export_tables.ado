/*
Input:
Models() nombre de los store de las regresiones previas

Output: ninguno en stata, pero exporta resultados similares a outreg2 a un excel, con soporte para hojas de cálculo

Opciones:
-Dec: número de decimales con los que se reportan los coeficientes
-Cell: Celda de excel en la que comienza la tabla
-Using: Nombre del archivo de Excel al cual se quiere exportar la tabla
-Sheet: Hoja de cálculo de Excel a la cual se quiere exportar la tabla
-Variables: Opcional. Si se especifica, se añade una columna con los nombres de las variables de los modelos. No se puede usar con Label (por defecto, no se muestra nada)
-Label: Opcional. Si se especifica, se añade una columna con los label de las variables de los modelos. No se puede usar con Variables (por defecto, no se muestra nada)
-Keep: Opcional. Si se usa, se muestran sólo las variables especificadas. No se puede usar con Drop (por defecto, se muestran todas las variables)
-Drop: Opcional. Si se usa, se muestran todas las variables menos las especificadas. No se puede usar con Keep (por defecto, se muestran todas las variables)
-Stats: Opcional. Añade al final de la tabla los estadísticos del tipo e() que se solicitan
-Std: Opcional. Añade los errores estándar abajo de los coeficientes. No se puede usar con Pvalues.
-Pvalues: Opcional. Añade los p-values abajo de los coeficientes. No se puede usar con Std.
-SEE: Opcional. Muestra un list con la tabla que se exportará a Excel

Pendiente: Añadir opción de cambiar los errores estándar por p-values
*/
capture program drop export_tables
program define export_tables, eclass
	version 15
	syntax , Models(string) Dec(real) Cell(string) USING(string) SHEET(string) [ Keep(string) Drop(string) DStats(real 0) STATS(string) VARiables LABEL STD PVALues  SEE ]
	quietly{
	*mensajes de error
	if "`var'"=="var" & "`label'"=="label" { //usar var y label al mismo tiempo
	noisily display as error "options label and var may not be combined"
	exit 184
	}
	if "`keep'"!="" & "`drop'"!="" { //usar keep y drop al mismo tiempo
	noisily display as error "options keep and drop may not be combined"
	exit 184
	}
	if "`std'"!="" & "`pvalues'"!="" { //usar keep y drop al mismo tiempo
	noisily display as error "options std and pvalues may not be combined"
	exit 184
	}
	
	
	preserve
	*pasamos los modelos para que se puedan usar en loop, y vemos cuantos son
	est table `models', keep(`keep') drop(`drop') stats(`stats')

	
	tempname index_temp //guardamos matrix para conocer posiciones de variables (modelos con distinta cantidad de variables)
	mat `index_temp'=r(coef)
	mat stats=r(stats)
	mat rownames stats=`stats'
	local num_models=colsof(r(coef))/2 //cantidad de modelos en tablelist
	local num_vars=rowsof(r(coef)) //cantidad de variables en tablelist
	
	local vars_temp="`:rownames r(coef)'" //nombres de las variables, para armar tabla
	*sacamos variables omitidas
	local vars ""
	foreach x in `vars_temp'{
	if !(substr("`x'",1,2)=="o."){
	local vars "`vars' `x'"
	}
	else{
	local --num_vars
	}
	}
	
	if "`label'"=="label"{
		foreach x in `vars'{
			if "`x'"=="_cons"{
				local `x' "Constant"
			}
			else{
				local `x': variable label `x'
			}
		}
	}
	
	*matrices para guardar datos finales
	tempname beta
	mat `beta'=J(`num_vars',`num_models',.)
	mat rownames `beta'=`vars' //pasamos los nombres de las columnas
	tempname pval
	mat `pval'=J(`num_vars',`num_models',.)
	mat rownames `pval'=`vars'
	tempname se
	mat `se'=J(`num_vars',`num_models',.)
	mat rownames `se'=`vars'
	
	*arreglamos matriz "index" para sacar las filas con variables omitidas
	
	tempname index
	local total=2*`num_models'
	mat `index'=J(`num_vars',`total',.)
	mat rownames `index'=`vars'
	forval k=1/`total'{
	local j=1
	foreach var in `vars'{
	mat `index'[`j',`k']=`index_temp'[rownumb(`index_temp',"`var'"),`k']
	local ++j
	}
	}

	local i=1
	foreach table in `r(names)'{ //loop sobre columnas de la tabla final
		est restore `table' //recuperamos la tabla
		eret display //vemos el display
		tempname temp //matriz temporal
		mat `temp'=r(table) //guardamos matriz
		foreach name in `vars'{ //loop para armar cada columna	
			capture mat `beta'[rownumb(`index',"`name'"),`i']=`temp'[rownumb(`temp',"b"),colnumb(`temp',"`name'")]
			capture mat `pval'[rownumb(`index',"`name'"),`i']=`temp'[rownumb(`temp',"pvalue"),colnumb(`temp',"`name'")]
			capture mat `se'[rownumb(`index',"`name'"),`i']=`temp'[rownumb(`temp',"se"),colnumb(`temp',"`name'")]
		}
		local ++i
	}	
	
	*ya tenemos las matrices necesarias, ahora sólo hay que juntarlas
	if "`std'"=="std" | "`pvalues'"=="pvalues"{
	if "`std'"=="std"{
	mat join= `beta' , `pval', `se'
	}
	if "`pvalues'"=="pvalues"{
	mat join= `beta' , `pval', `pval'
	}
	mat li join
	clear //limpiamos la base, para guardar aquí las tablas
	svmat join
	local stub=""
	forval x=1/`num_models'{
		local y=`num_models' + `x'
		local z=2*`num_models' + `x'
		rename (join`x' join`z' join`y') (beta`x'1 beta`x'2 pval`x'1)
		gen pval`x'2=.
		local stub "`stub' beta`x' pval`x'"
	}
	gen id=_n
	reshape long `stub', i(id) j(j)
	drop id
	gen par1="("
	gen par2=")"
	forval x=1/`num_models'{
		gen stars`x'="" //pasamos p-values a estrellas
		replace stars`x'="*" if pval`x'<0.1
		replace stars`x'="**" if pval`x'<0.05
		replace stars`x'="***" if pval`x'<0.01
		gen temp`x'= string(beta`x', "%10.`dec'f") //truncamos decimales coeficientes
		egen beta_`x'=concat( temp`x' stars`x') //juntamos coeficientes con estrellas
		if "`std'"=="std"{
			egen sub_`x'=concat(par1 temp`x' par2) //le ponemos paréntesis a los SE
		}
		if "`pvalues'"=="pvalues"{
			gen sub_`x'=temp`x' // pasamos los p-values
		}
		gen var`x'=""
		replace var`x'=beta_`x' if j==1
		replace var`x'=sub_`x' if j==2
		replace var`x'="-" if pval`x'==. & j==1
		replace var`x'="-" if var`x'=="(.)" | var`x'=="."
	}
	}
	else {
	mat join= `beta' , `pval'
	mat li join
	clear //limpiamos la base, para guardar aquí las tablas
	svmat join
	di `num_models'
	forval x=1/`num_models'{
		local y=`num_models' + `x'
		gen stars`x'="" //pasamos p-values a estrellas
		replace stars`x'="*" if join`y'<0.1
		replace stars`x'="**" if join`y'<0.05
		replace stars`x'="***" if join`y'<0.01
		gen temp`x'= string(join`x', "%10.`dec'f") //truncamos decimales coeficientes
		egen var`x'=concat( temp`x' stars`x') //juntamos coeficientes con estrellas
		replace var`x'="-" if join`y'==.
	}
	}
	keep var* //guardamos lo relevante
	if "`label'"=="label"{  //nombres de los label de variables
		gen label=""
		local j=1
		foreach x in `vars'{
			replace label="``x''" in `j'
			local ++j
			if "`std'"=="std" | "`pvalues'"=="pvalues"{
			local ++j
			}
		}
	order label
	}
	if "`variables'"=="variables"{ //nombres de las variables
		gen name=""
		local j=1
		foreach x in `vars'{
			if "`x'"=="_cons"{
				replace name="constant" in `j'
			}
			else {
				replace name="`x'" in `j'
			}
			local ++j
			if ("`std'"=="std"  | "`pvalues'"=="pvalues") local ++j
		}
	order name
	}
	if "`stats'"!=""{
		*guardamos base previa
		tempfile main
		save `main'
		clear
		*trabajamos base con los stats
		svmat stats
		forval x=1/`num_models'{
		if "`dstats'"!="."{
		gen var`x'= string(stats`x', "%10.`dstats'f")
		}
		else{
		gen var`x'= string(stats`x', "%10.`dec'f")
		}
		replace var`x'="-" if var`x'==".z"
		}
		*añadimos nombres de los estadísticos
		if "`label'"=="label"{
		local name "label"
		}
		else if "`variables'"=="variables"{
		local name "name"
		}
		if "`name'"!=""{
		gen `name'=""
		local m=1
		foreach stat in `stats'{
		replace `name'="`stat'" in `m'
		local ++m
		}
		}
		*le ponemos 0 decimales al N
		forval x=1/`num_models'{
		replace var`x'=string(stats`x', "%10.0f") if `name'=="N"
		}
		keep var* `name'
		tempfile base2
		save `base2'
		*pegamos ambas bases
		use `main', clear
		append using `base2'
	}
	
	*sacamos filas vacías (omitidas en todos los modelos)
	if ("`std'"=="std" | "`pvalues'"=="pvalues")  local num_vars= 2 * `num_vars'
	
	
	gen cond=0
	replace cond=1 in 1/`num_vars'
	forval i=1/`num_vars'{
	forval j=1/`num_models'{
	if var`j'[`i']!="-" replace cond=0 in `i'
	}
	}
	drop if cond==1
	drop cond
	
	
	noisily export excel using `using', sheet(`sheet') sheetmodify cell(`cell')
	matrix drop join stats
	if "`see'"=="see" no list
	restore
}
end
