# Sistema-Automatico-Para-Estudiantes
Sistema basado en la nube empleando Google Apps Script para registro de estudiantes, verificación difusa y sincronización con equipo de mentores

# 🎓 Sistema de Automatización de Datos Estudiantiles (Google Apps Script)

Arquitectura de datos escalable diseñada para gestionar +3,000 registros de estudiantes distribuidos en un ecosistema de 30+ hojas de cálculo independientes para mentores.

# Funcionalidades Clave
* **Lógica de Coincidencia Difusa (Fuzzy Matching):** Implementación del algoritmo de Distancia de Levenshtein para detectar y marcar discrepancias en CURP o nombres (activación de revisión manual).
* **Arquitectura Push-Pull:** Sincronización masiva entre un documento Maestro y hojas de mentores mediante ejecución en servidor, optimizando el rendimiento al eliminar fórmulas pesadas de `IMPORTRANGE`.
* **Consolidación Automática de Datos:** Agregación de registros desde múltiples formularios externos en una pestaña unificada con rastreo de origen (`ACT. DATOS` vs `NVOS. MB`).
* **Integridad de Datos Visual:** Sistema de alertas automáticas (Naranja/Amarillo) para inconsistencias en niveles educativos o discrepancias de grados escolares.

## 🛠️ Stack Tecnológico
* **Lenguaje:** JavaScript (Google Apps Script)
* **Algoritmos:** Distancia de Levenshtein (Similitud de Cadenas)
* **Entorno:** Google Workspace API / Triggers basados en tiempo

## 📁 Estructura del Repositorio
* `MENU.gs`: Menu central ejecutado al abrir el proyecto y funciones adicionales de normalización de texto.
* `PUSH-PULL 2.gs`: Script central para la vinculación del sistema maestro con las listas individualizadas de +30 mentores garantizando la integridad de los datos mediante un sistema PUSH-PULL con archivo historico de estudiantes traspasados.
* `FORMS-STATUS.gs`: Lógica central para el estatus de registro de alumnos y alertas por correo.
* `LISTA-FORMS.gs`: Funciones para la unión de bases de datos externas y limpieza de información.
* `FORMATO CONDICIONAL.gs`: Script optimizado de `onEdit` para la gestión visual de niveles y estados.
* `COPIA DE SEGURIDAD.gs`: Creación de un copias de seguridad estaticas periodicas del documento master.
