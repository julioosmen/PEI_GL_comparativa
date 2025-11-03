# PEI_GL_comparativa
El objetivo general es desarrollar un aplicativo/sistema que permita comparar y verificar tablas de OEI/AEI con la matriz estándar a nivel de gobierno local

* Actualización al 02.11.2025
* MEJORAS:
*   i) Los códigos "compare_oei" y "compare_aei" se han modificado para que la comparación que realizan en las funciones definidas no consideren comas, tildes, espacios. Para ello se ha normalizan los textos para que la comparación se concentre en las palabras incluidas en las frases comparadas.
*   ii) Se ha eliminado la columna "similitud" en los resultados mostrados en el streamlit dado que no suma que se muestre un valor de la comparación. Lo que se agregó es una columna de "diferencias" que muestra las diferencias literales de la comparación. Con esta columna se apoya al especialista para que identifique fácilmente las comparaciones. 
