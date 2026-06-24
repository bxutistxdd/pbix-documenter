// Known Azure DevOps OData field descriptions
export const ODATA_FIELD_DESCRIPTIONS = {
  "WorkItemId": "Identificador único del work item en Azure DevOps",
  "WorkItemType": "Tipo de work item (Epic, Feature, User Story, Bug, Task, etc.)",
  "Title": "Título o nombre descriptivo del work item",
  "State": "Estado actual del work item en el flujo de trabajo (New, Active, Resolved, Closed, etc.)",
  "StateCategory": "Categoría del estado (Proposed, InProgress, Resolved, Completed)",
  "Priority": "Prioridad asignada al work item (1=Crítica, 2=Alta, 3=Media, 4=Baja)",
  "Severity": "Nivel de severidad del defecto o incidente",
  "AssignedTo": "Usuario o equipo asignado para trabajar en el work item",
  "AssignedToUserSK": "Clave surrogate del usuario asignado",
  "AreaSK": "Clave surrogate del área de iteración",
  "AreaId": "Identificador del área de iteración",
  "AreaPath": "Ruta jerárquica del área dentro del proyecto Azure DevOps",
  "IterationSK": "Clave surrogate de la iteración (sprint)",
  "IterationId": "Identificador de la iteración (sprint)",
  "IterationPath": "Ruta jerárquica de la iteración dentro del proyecto",
  "ChangedDate": "Fecha y hora de la última modificación del work item",
  "ChangedDateSK": "Clave surrogate de la fecha de última modificación",
  "CreatedDate": "Fecha y hora de creación del work item",
  "CreatedDateSK": "Clave surrogate de la fecha de creación",
  "ClosedDate": "Fecha y hora en que el work item fue cerrado",
  "ClosedDateSK": "Clave surrogate de la fecha de cierre",
  "ResolvedDate": "Fecha en que el work item fue marcado como resuelto",
  "ResolvedDateSK": "Clave surrogate de la fecha de resolución",
  "ActivatedDate": "Fecha en que el work item fue activado o iniciado",
  "StateChangeDate": "Fecha del último cambio de estado",
  "StateChangeDateSK": "Clave surrogate de la fecha de cambio de estado",
  "CommentCount": "Número de comentarios registrados en el work item",
  "IsActive": "Indica si el work item está actualmente activo (true/false)",
  "IsLastRevisionOfDay": "Indica si es la última revisión del work item en ese día",
  "IsLastRevisionOfPeriod": "Indica si es la última revisión del período analizado",
  "ParentWorkItemId": "ID del work item padre en la jerarquía",
  "TagNames": "Etiquetas (tags) asignadas al work item",
  "TeamProject": "Nombre del proyecto de equipo en Azure DevOps",
  "Watermark": "Marca de agua para control de sincronización incremental",
  "StoryPoints": "Puntos de historia estimados para el work item",
  "Effort": "Esfuerzo estimado en horas o puntos",
  "RemainingWork": "Trabajo restante estimado",
  "OriginalEstimate": "Estimación original de trabajo",
  "CompletedWork": "Trabajo completado registrado",
  "BusinessValue": "Valor de negocio asignado al work item",
  "TimeCriticality": "Criticidad temporal del work item",
  "Risk": "Nivel de riesgo asociado",
  "Blocked": "Indica si el work item está bloqueado",
  "ResolvedBy": "Usuario que resolvió el work item",
  "ClosedBy": "Usuario que cerró el work item",
  "CreatedBy": "Usuario que creó el work item",
  "ChangedBy": "Usuario que realizó el último cambio",
  "WorkItemRevisionSK": "Clave surrogate de la revisión del work item",
  "DateSK": "Clave surrogate de fecha para relación con tabla calendario",
  "Date": "Fecha de referencia del snapshot o registro",
  "Count": "Conteo de registros o work items",
  "LeadTimeDays": "Días de tiempo de entrega (lead time) desde creación hasta cierre",
  "CycleTimeDays": "Días de tiempo de ciclo desde activación hasta cierre",
  "InProgressDate": "Fecha en que el work item pasó a estado En Progreso",
  "TestSuiteId": "Identificador único del suite de pruebas",
  "TestSuiteName": "Nombre del suite de pruebas",
  "TestPlanId": "Identificador único del plan de pruebas",
  "TestPlanName": "Nombre del plan de pruebas",
  "TestCaseId": "Identificador único del caso de prueba",
  "TestCaseName": "Nombre o título del caso de prueba",
  "TestCaseTitle": "Título descriptivo del caso de prueba",
  "TestResultId": "Identificador del resultado de ejecución",
  "Outcome": "Resultado de la ejecución del caso de prueba (Passed, Failed, Blocked, etc.)",
  "OutcomeLastUpdatedDate": "Fecha de la última actualización del resultado",
  "RunBy": "Usuario que ejecutó la prueba",
  "Configuration": "Configuración bajo la cual se ejecutó la prueba",
  "TestSuiteType": "Tipo de suite (Static, DynamicByQuery, DynamicByRequirement)",
  "ProjectSK": "Clave surrogate del proyecto",
  "AnalyticsUpdatedDate": "Fecha de última actualización en el servicio Analytics de Azure DevOps",
};

// Known table context descriptions
export const TABLE_CONTEXT = {
  "FechasHabiles": "Lista de SharePoint con las fechas hábiles del año. Se usa para calcular días hábiles transcurridos excluyendo fines de semana y festivos.",
  "Fechas Habiles 2026": "Lista de SharePoint con las fechas hábiles del año 2026. Se usa para calcular días hábiles transcurridos excluyendo fines de semana y festivos.",
  "HierarchyFlat": "Tabla derivada de las consultas OData de Azure DevOps que representa la jerarquía a tiempo real de cada work item (Epic → Feature → User Story → Bug). Permite navegar la estructura jerárquica del proyecto.",
  "HierarchyTest": "Tabla derivada de las consultas OData que contiene la jerarquía de work items relacionados con testing: TestSuites, TestPlans y demás elementos del apartado de pruebas.",
  "PApp_Context": "Tabla de texto plano con una serie de filas que sirven como contexto para la PowerApp embebida en el reporte, permitiendo que sepa en qué página del reporte se encuentra el usuario actualmente.",
};

export function getColumnDescription(tableName, colName, aiDescription) {
  if (aiDescription && aiDescription.length > 3 && aiDescription !== "—") return aiDescription;
  // Check OData known fields
  const known = ODATA_FIELD_DESCRIPTIONS[colName];
  if (known) return known;
  // Partial match
  for (const [key, desc] of Object.entries(ODATA_FIELD_DESCRIPTIONS)) {
    if (colName.toLowerCase().includes(key.toLowerCase()) || key.toLowerCase().includes(colName.toLowerCase())) {
      return desc;
    }
  }
  return aiDescription || "";
}

export function getTableContext(tableName, aiDescription) {
  for (const [key, desc] of Object.entries(TABLE_CONTEXT)) {
    if (tableName.toLowerCase().includes(key.toLowerCase())) return desc;
  }
  return aiDescription || "";
}
