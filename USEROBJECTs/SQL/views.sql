CREATE VIEW SEI_MONITORIZEDFIELDS
AS
SELECT T0.Code, T0.U_FormTypeEx as FormTypeEx, T1.U_Tabla AS Tabla, T1.U_CampoBBDD AS CampoBD, T1.U_Objeto AS CampoObjeto,T1.U_Especial AS EsEspecial,ISNULL(T1.U_TablaAsoc,'') AS TablaAsociada, ISNULL(T1.U_CampoAsoc,'') AS CampoAsociado,  T1.U_Descripcion  AS Description
FROM [@SEI_INTERFAZC] AS T0
INNER JOIN [@SEI_INTERFAZL] AS T1 ON T0.Code = T1.Code
$
