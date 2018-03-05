'*******************************************************************************************************************************************************
'* NOMBRE DE LA FUNCIÓN: Reto_Eduard_Correa                                                
'* DESCRIPCIÓN: Dando solución al ejercicio propuesto, se crea este programa que permite capturar las carácteristicas de la tarjeta 
'*              American Express y MastercadrBlack. También cuenta con la capacidad de diligenciar el formulario para Amex de forma paramétrica. 
'*
'* PARÁMETROS DE ENTRADA: N/A
'*
'* PARÁMETROS DE SALIDA: N/A
'*	
'*
'* NOTA: Se trabaja con el navegador GOOGLE CHROME en su Versión 64.0.3282.186 (Build oficial) (64 bits)
'*
'* Autor: Eduard Correa
'* Fecha Creación: 03/03/2018'*
'*******************************************************************************************************************************************************

'Declaración de Variables.
	Dim DataDriven, Repositorio
	Dim Url, Urlsoporte, Repo

'Se parametriza el Datadriven para que sea dinámico
	DataDrivenPath = "\Documents\Reto_Eudard_1.0\REPOSITORIO.xls"
	DataDriven = CreateObject ("WScript.Shell").ExpandEnvironmentStrings("%USERPROFILE%") & DataDrivenPath

'Se deja en una variable la URL con la que se va a trabajar
	Url = "https://www.grupobancolombia.com/wps/portal/personas"


'Importo el DataTable
	DataTable.Import (DataDriven)
	
'Se abre el navegador Google Chrome
	SystemUtil.Run "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
	
'1. Ingreso a la opción productos y servicios y listo todas las opciones mostradas.
	Browser("Browser").Navigate (Url)
	Browser("Browser").Page("Personas: Soluciones Financier").Link("Productos y Servicios").Click
		
'2.	Ingreso a la opción Tarjetas de Crédito. 	
	Browser("Browser").Page("Personas: Soluciones Financier").Link("Tarjetas Crédito").Click
	
'3. Llevo a un archivo .xls la información completa de la tarjeta de crédito American Express y MasterCard Black.	
	Call TomarValoresyllevarExcel()

'4. Se da clic en el botón “Solicítala aquí” de la tarjeta American Express. 
	Browser("Browser").Page("Tarjeta de Crédito para").Link("Solicítala aquí").Click

'Controlamos que el formulario esté disponible para empezar a diligenciar los datos.
	Browser("Browser").Page("Solicitud Tarjeta de Crédito").Sync	
	While Browser("Browser").Page("Solicitud Tarjeta de Crédito").Frame("Frame_3").WebElement("Cargando En este momento").GetROProperty("Visible") = True 
	wait(1)
	Wend

'5. Se procede con el llenado del formulario, trayendo los datos desde el Datadriven.
	Call LlenadoDatos()	
	
'Damos Clic en el botón continuar 	
	Browser("Browser").Page("Solicitud Tarjeta de Crédito").Frame("Frame_2").WebButton("Continuar").Click
	wait(2)
		
'Se cierra el Navegador
	Browser("Browser").Close

'Se exporta el DataDriven.
	DataTable.Export(DataDriven)
