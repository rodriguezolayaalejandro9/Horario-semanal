# Horario-semanal

# Generador Automático de Horario Semanal

## 📌 Lo que hace
Este repositorio contiene dos scripts que generan automáticamente el horario semanal para los estudiantes del Colegio Gimnasio Kaiporé:

1. **Primaria (1° a 5°)**  
   - Identifica áreas con menor avance por estudiante.  
   - Completa los huecos en el horario priorizando las asignaturas con menor porcentaje de cumplimiento.  
   - Respeta los cupos máximos por área.  
   - Exporta el horario final a un archivo Excel listo para su uso.

2. **Bachillerato (6° a 11°)**  
   - Aplica la misma lógica adaptada a las áreas y asignaturas específicas de bachillerato.  
   - Maneja cupos y prioridades según las necesidades de este nivel académico.

---

## 📚 Librerías usadas
- pandas  
- numpy  

Además, requiere los módulos personalizados:
- `corregir_nombres`  

---

## 📝 Notas
- Este proyecto ahorró aproximadamente **5 horas de trabajo semanal** de la coordinadora general del colegio, quien antes se encargaba manualmente de elaborar el horario semanal.
- Está adaptado a la estructura y necesidades específicas del Colegio Gimnasio Kaiporé, con configuraciones distintas para primaria y bachillerato.
