import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import * as mammoth from 'mammoth';
import {
  Upload,
  FileText,
  Sparkles,
  ChevronRight,
  ChevronLeft,
  Download,
  Plus,
  Trash2,
  Edit3,
  Check,
  X,
  Loader2,
  Archive,
  Home,
  Building2,
  User,
  Palette,
  AlertCircle,
  RefreshCw,
  Eye,
  Save,
  Paperclip,
  File as FileIcon,
  MapPin,
  Globe,
  TrendingUp,
} from 'lucide-react';
import jsPDF from 'jspdf';
import 'jspdf-autotable';

// ============================================================
// LOGO EVALUAR — Wordmark SVG (siempre se ve bien, cualquier tamaño)
// ============================================================
// Componente que renderiza el wordmark "evaluar.ai" con los colores de marca.
// Uso: <EvaluarLogo className="h-9" />
const EvaluarLogo = ({ className = '', opacity = 1, ...props }) => (
  <svg
    viewBox="0 0 190 44"
    xmlns="http://www.w3.org/2000/svg"
    className={className}
    style={{ opacity }}
    aria-label="evaluar.ai"
    role="img"
    {...props}
  >
    <text
      x="0"
      y="32"
      fontFamily="'Manrope', sans-serif"
      fontSize="32"
      fontWeight="800"
      fill="#2E1C5A"
      letterSpacing="-1.2"
    >
      evaluar
    </text>
    <circle cx="142" cy="28" r="4.5" fill="#D91B7E" />
    <text
      x="152"
      y="32"
      fontFamily="'Manrope', sans-serif"
      fontSize="16"
      fontWeight="600"
      fill="#6B6480"
      letterSpacing="-0.4"
    >
      ai
    </text>
  </svg>
);

// ============================================================
// DICCIONARIO EVALUAR — 149 competencias en 12 categorías
// Fuente: Catálogo Evaluar Marzo 2026
// ============================================================
const DICCIONARIO_EVALUAR = {
  Actitudinales: {
    'Alta Energía y dinamismo':
      'Impulso constante para mantener un ritmo activo de trabajo, mostrando entusiasmo e iniciativa para responder con agilidad a las demandas y sostener dinamismo en distintos contextos.',
    'Apertura al cambio':
      'Preparación para incorporar nuevas ideas y adaptarse con rapidez a situaciones emergentes, ajustando métodos y explorando alternativas que permiten asumir desafíos y favorecer la evolución del trabajo.',
    'Apertura mental':
      'Interés por considerar ideas y perspectivas diversas, ajustando el pensamiento con flexibilidad e incorporando nueva información para facilitar el aprendizaje y la adaptación.',
    'Cosmopolitismo/Visión internacional':
      'Facilidad para desenvolverse en contextos culturales diversos, integrando perspectivas globales y adaptando el trabajo a realidades internacionales.',
    Disponibilidad:
      'Flexibilidad para atender requerimientos y ofrecer apoyo cuando es necesario, adaptándose a cambios de ritmo, necesidades o prioridades para asegurar continuidad y buen servicio.',
    'Flexibilidad y Adaptabilidad':
      'Apertura para ajustarse con rapidez a nuevos escenarios, demandas o formas de trabajo, modificando conductas y pensamientos para mantener eficacia y continuidad en contextos cambiantes.',
    'Independencia/autonomía':
      'Habilidad para avanzar por cuenta propia, tomando decisiones responsables y organizando el trabajo sin depender de supervisión constante, garantizando continuidad y calidad en los resultados.',
    Iniciativa:
      'Impulso para actuar antes de que las situaciones lo exijan, explorando alternativas y promoviendo acciones que aportan soluciones y favorecen el avance del trabajo sin depender de instrucciones directas.',
    Innovación:
      'Impulso para explorar ideas novedosas y proponer alternativas distintas que mejoren procesos o resuelvan problemas, integrando creatividad e iniciativa para convertir desafíos en oportunidades.',
    'Life balance/Equilibrio trabajo-vida':
      'Capacidad de equilibrar la vida laboral y personal, encontrando el espacio adecuado para cada momento, solucionando los conflictos que se puedan producir entre ambos.',
    'Mejora continua':
      'Interés permanente por elevar la calidad del trabajo, identificando oportunidades de optimización y ajustando métodos o procesos para avanzar hacia estándares superiores de desempeño.',
    'Perseverancia y disciplina':
      'Constancia para sostener el esfuerzo a lo largo del tiempo, manteniendo orden y compromiso frente a las dificultades para asegurar avances estables y confiables hacia las metas.',
    Proactividad:
      'Iniciativa constante para anticiparse a necesidades y oportunidades, impulsando acciones y propuestas que contribuyen al logro de metas y generan valor sin esperar indicaciones.',
    'Tesón y Empeño':
      'Dedicación constante para avanzar hacia los objetivos, manteniendo disciplina, cuidado en los detalles y determinación frente a las dificultades.',
    'Toma de riesgos y decisiones':
      'Facilidad para decidir con criterio en contextos inciertos, evaluando alternativas y riesgos para actuar con oportunidad y avanzar en situaciones que requieren determinación.',
    Urgencia:
      'Reacción ágil ante demandas inmediatas, conservando claridad mental para priorizar con rapidez y ejecutar tareas críticas con dinamismo y precisión incluso bajo presión.',
    'Interés por la Estabilidad (Job stability)':
      'Capacidad de alinear los intereses personales con los objetivos de la organización, demostrando lealtad, estabilidad y contribución al crecimiento sostenido.',
    'Proyección Profesional (Reskilling)':
      'Capacidad de transformarse y adquirir nuevas competencias ante cambios en el entorno, manteniéndose relevante y aportando valor de forma continua.',
  },
  Administrativas: {
    'Administración de procesos':
      'Capacidad para analizar, organizar y optimizar procesos de manera estructurada, comprendiendo la relación entre actividades y ajustando los flujos de trabajo para asegurar eficiencia, coherencia y continuidad operativa.',
    'Administración del tiempo':
      'Capacidad para organizar tareas y prioridades con disciplina, planificando el trabajo con criterios lógicos y manteniendo el enfoque necesario para cumplir plazos y objetivos con eficiencia.',
    'Calidad del trabajo':
      'Enfoque constante en la precisión y el cuidado por los detalles, manteniendo estándares altos y una búsqueda continua de mejorar la forma en que se ejecutan las tareas.',
    'Capacidad de gestión':
      'Capacidad para coordinar tareas, recursos y decisiones de manera efectiva, asegurando el avance ordenado de proyectos y el cumplimiento oportuno de los objetivos.',
    'Capacidad de planificación':
      'Capacidad para ordenar información, recursos y tiempos de manera anticipada, estructurando planes que conectan prioridades con objetivos y permiten avanzar con claridad hacia el cumplimiento de metas.',
    'Conciencia organizacional':
      'Comprensión del funcionamiento, la cultura y las dinámicas internas de la organización, alineando las acciones personales con el propósito, los valores y las prioridades colectivas.',
    'Escrupulosidad/Minuciosidad':
      'Disposición a trabajar con un alto nivel de detalle y rigor, asegurando exactitud y consistencia en cada tarea, incluso cuando los procesos requieren precisión sostenida.',
    'Orden y calidad (QA)':
      'Enfoque sistemático para mantener procesos y entregables organizados, coherentes y alineados con estándares de calidad, incluso en contextos dinámicos o digitales.',
    'Pensamiento estratégico':
      'Capacidad para analizar información de manera amplia, anticipar escenarios y orientar decisiones que conecten las acciones presentes con resultados sostenibles en el largo plazo.',
    Productividad:
      'Habilidad para sostener un ritmo de trabajo constante y enfocado, utilizando el tiempo y la energía con eficiencia para avanzar con dinamismo y aportar al cumplimiento de los resultados.',
    'Razonamiento financiero':
      'Capacidad para analizar información numérica y económica, interpretar tendencias y evaluar escenarios que permitan tomar decisiones financieras acertadas para el negocio.',
    'Resultados Ágiles (Results Agility)':
      'Capacidad de alcanzar objetivos sobresalientes en entornos cambiantes, adaptándose rápidamente, tomando decisiones bajo presión y encontrando soluciones innovadoras.',
    'Optimización de Recursos':
      'Habilidad para organizar y ejecutar tareas utilizando eficientemente los recursos disponibles, garantizando resultados con el menor costo y esfuerzo posibles.',
    'Trabajo sin supervisión':
      'Capacidad de planificar, organizar y ejecutar actividades de manera autónoma, priorizando eficientemente el trabajo para cumplir objetivos sin necesidad de supervisión constante.',
  },
  Cognitivas: {
    'Actualización del Conocimiento':
      'Interés permanente por aprender y mantenerse al día, incorporando información reciente con criterio y aplicando nuevos contenidos para mejorar el desempeño.',
    'Agilidad mental (Mental Agility)':
      'Capacidad para procesar información con rapidez y precisión, detectando patrones y adaptándose a demandas cambiantes para responder con claridad y tomar decisiones efectivas.',
    'Aptitud mecánica':
      'Facilidad para comprender el funcionamiento de mecanismos y sistemas físicos, interpretando relaciones entre piezas, fuerzas y movimientos para anticipar fallas y resolver problemas técnicos.',
    'Aptitud numérica':
      'Habilidad para interpretar y trabajar con datos numéricos con precisión, identificando relaciones y patrones que facilitan el análisis de información y la toma de decisiones.',
    'Aptitud espacial':
      'Habilidad para interpretar y visualizar relaciones entre objetos, distancias y estructuras, comprendiendo representaciones gráficas o tridimensionales con precisión.',
    'Aptitudes verbales':
      'Habilidad para comprender y expresar ideas con claridad, utilizando el lenguaje de manera precisa y estructurada para interpretar información y comunicarse con fluidez.',
    'Construcción del conocimiento en equipo':
      'Participación activa en la creación y consolidación de conocimiento colectivo, aportando análisis que clarifican información.',
    Creatividad:
      'Aptitud para generar ideas originales y explorar nuevas posibilidades, integrando perspectivas diversas para convertirlas en soluciones prácticas o innovadoras.',
    'Discernimiento y criterio':
      'Capacidad para evaluar información con rigor, distinguir lo relevante y tomar decisiones fundamentadas aplicando un juicio equilibrado que anticipa consecuencias.',
    'Pensamiento Abstracto':
      'Aptitud para interpretar ideas, relaciones y estructuras no literales, analizando información de manera flexible para comprender conceptos complejos.',
    'Pensamiento Analítico':
      'Capacidad para examinar información de manera lógica y estructurada, descomponiendo problemas para identificar patrones, causas y posibles soluciones.',
    'Pensamiento Conceptual':
      'Capacidad para interpretar conceptos, relaciones y patrones que explican ideas complejas, organizando el pensamiento de forma estructurada.',
    'Rapidez y exactitud perceptiva':
      'Aptitud para identificar detalles, cambios o patrones con rapidez y precisión, manteniendo atención sostenida.',
    'Resolución de problemas':
      'Capacidad para comprender situaciones, analizar información con lógica y definir soluciones prácticas que respondan con claridad y criterio a los retos.',
    'Gestión del Saber Corporativo':
      'Capacidad de adquirir, organizar y compartir el conocimiento generado dentro de la organización, optimizando la toma de decisiones y la innovación.',
    'Pensamiento Secuencial':
      'Capacidad de analizar y estructurar actividades de manera lógica y lineal, visualizando su desarrollo en una secuencia coherente.',
    'Mentalidad Práctica':
      'Capacidad de abordar problemas y situaciones con un enfoque realista y orientado a la acción, priorizando soluciones factibles y eficaces.',
    'Análisis de Datos (Data Analysis)':
      'Capacidad de analizar grandes volúmenes de información para reconocer, interpretar y anticipar comportamientos o tendencias recurrentes.',
    'Aprendizaje Ágil (Upskilling)':
      'Capacidad de adquirir y aplicar nuevos conocimientos con rapidez y efectividad, ajustándose con facilidad a entornos cambiantes.',
    'Inteligencia General':
      'Capacidad para comprender información con rapidez, integrar conceptos con facilidad y resolver problemas mediante un razonamiento lógico y flexible.',
  },
  Comerciales: {
    'Agresividad comercial (tipo cazador)':
      'Capacidad para generar nuevas oportunidades de negocio mediante iniciativa, determinación y una fuerte orientación a resultados, enfocándose en abrir mercados.',
    'Credibilidad Comercial':
      'Disposición a generar confianza y ser percibido como una figura confiable y profesional, construyendo relaciones comerciales sólidas.',
    'Desarrollo de redes comerciales':
      'Estilo orientado a crear y ampliar vínculos que generan oportunidades de negocio, relacionándose con seguridad y dinamismo.',
    Mediación:
      'Disposición a facilitar el entendimiento entre partes con intereses distintos, abordando los desacuerdos con objetividad.',
    'Negociación efectiva':
      'Habilidad para construir acuerdos equilibrados mediante una interacción estratégica y orientada a intereses, manteniendo claridad y enfoque en resultados.',
    'Orientación al Cliente':
      'Actitud genuina de comprender y atender al cliente, priorizando sus necesidades y buscando generar una experiencia positiva y de confianza en cada interacción.',
    'Orientación y adaptabilidad a las ventas':
      'Disposición flexible para desenvolverse con eficacia en distintos entornos comerciales, ajustando el estilo y el enfoque para conectar con el cliente.',
    'Perseverancia comercial (tipo agricultor)':
      'Enfoque constante y orientado al largo plazo para sostener y fortalecer relaciones comerciales, impulsando el crecimiento continuo de las cuentas.',
    'Visión del Negocio':
      'Capacidad para interpretar el entorno y conectar esa información con los objetivos organizacionales, anticipando necesidades y orientando decisiones estratégicas.',
  },
  Comunicacionales: {
    Asertividad:
      'Claridad respetuosa al comunicar ideas, necesidades y límites, permitiendo dialogar con apertura y mantener conversaciones directas incluso ante diferencias.',
    'Comunicación efectiva':
      'Capacidad para transmitir ideas con claridad y coherencia, ajustando el mensaje al contexto y asegurando una comprensión precisa.',
    'Escucha activa':
      'Forma de atender y comprender de manera íntegra a la otra persona, interpretando su mensaje y contexto emocional para asegurar un entendimiento claro.',
    'Habilidades de contacto':
      'Facilidad para iniciar y sostener interacciones que generan cercanía y confianza, comunicándose con apertura y asertividad.',
    'Habilidades de Redacción':
      'Capacidad para expresar ideas de manera clara, precisa y coherente, estructurando mensajes bien organizados y adaptados al contexto.',
    'Habilidades mediáticas':
      'Aptitud para transmitir ideas con claridad y presencia en distintos medios, ajustando el mensaje y el estilo comunicativo para generar impacto.',
    'Impacto e influencia':
      'Potencial para comunicar y sostener ideas con claridad y presencia, logrando persuadir, movilizar y generar apoyo para avanzar hacia objetivos compartidos.',
    'Relaciones públicas':
      'Función orientada a representar a la organización ante distintos públicos, comunicando con coherencia y fortaleciendo su imagen.',
    'English Communication Skills':
      'Capacidad de comprender, interpretar y utilizar el idioma inglés de manera efectiva en diversos contextos.',
  },
  Digital: {
    'Búsqueda de datos e información':
      'Capacidad para identificar, explorar y verificar fuentes relevantes, utilizando criterios analíticos y herramientas digitales para sustentar decisiones.',
    'Comunicación Digital':
      'Destreza para transmitir ideas con claridad y eficacia en entornos virtuales, adaptando el mensaje al medio y utilizando herramientas tecnológicas con soltura.',
    'Cultura Digital (Habilidad Digital)':
      'Dominio práctico de herramientas y entornos digitales, comprendiendo su funcionamiento y aplicando criterios seguros para proteger la información.',
    'E-people/Ciudadano Digital':
      'Actitud responsable y activa en el uso de entornos y herramientas digitales, adoptando prácticas seguras y conscientes que favorecen la colaboración.',
    'Networking/Relaciones Digitales':
      'Forma de construir y sostener vínculos profesionales en entornos digitales, comunicándose con claridad y utilizando herramientas tecnológicas.',
    Teletrabajo:
      'Autonomía para organizarse y mantener desempeño efectivo en entornos virtuales, gestionando tareas y metas con claridad y utilizando herramientas digitales.',
  },
  Educativas: {
    'Aprendizaje continuo':
      'Interés sostenido por aprender y actualizarse, incorporando nuevas ideas y experiencias que enriquecen el desempeño y favorecen el desarrollo.',
    'Búsqueda de conocimientos':
      'Interés constante por ampliar la comprensión a través de nueva información y experiencias, integrando análisis y diversas perspectivas.',
    Didáctica:
      'Claridad para transmitir conocimientos de manera dinámica y accesible, adaptándose a distintos estilos de aprendizaje.',
    'Orientación/Asesoramiento':
      'Facultad para ofrecer guía clara y apoyo efectivo, integrando criterio analítico y comunicación precisa para resolver dudas y facilitar la comprensión.',
    'Mentoría efectiva':
      'Capacidad de guiar y desarrollar el potencial de otros mediante la transmisión de conocimientos, acompañamiento emocional y retroalimentación constructiva.',
  },
  Emocionales: {
    Afabilidad:
      'Trato amable y respetuoso que favorece relaciones cercanas, integrando sensibilidad interpersonal y apertura a distintas perspectivas para generar confianza.',
    Autocontrol:
      'Estabilidad emocional para regular reacciones y mantener claridad incluso bajo presión, gestionando impulsos y respondiendo con calma.',
    Cautela:
      'Prudencia para evaluar riesgos y actuar con autocontrol, respetando normas y evitando respuestas impulsivas para asegurar un desempeño responsable y seguro.',
    'Confianza en sí mismo':
      'Seguridad personal para sostener decisiones y enfrentar desafíos, reconociendo fortalezas y límites y proyectando autocontrol.',
    'Inteligencia emocional':
      'Capacidad para reconocer y regular las propias emociones, manteniendo estabilidad aun bajo presión, y para comprender a los demás.',
    Optimismo:
      'Visión constructiva frente a desafíos y cambios, interpretando las situaciones con esperanza realista y confianza para avanzar.',
    Resiliencia:
      'Fortaleza para recuperarse ante dificultades, conservando estabilidad emocional, claridad para evaluar la situación y la determinación para reajustar.',
    'Temple y Serenidad':
      'Calma sostenida para enfrentar presión o incertidumbre, respondiendo con equilibrio emocional, claridad y objetividad.',
    'Trabajo bajo presión':
      'Rendimiento consistente en escenarios de alta demanda, gestionando las exigencias con claridad mental y equilibrio emocional.',
  },
  Relacionales: {
    'Apoyo a los compañeros':
      'Actitud colaborativa que impulsa el bienestar y el desempeño colectivo, ofreciendo ayuda oportuna y generando un clima de confianza.',
    Colaboración:
      'Actitud orientada al trabajo conjunto, aportando de manera abierta y respetuosa para facilitar la comunicación, apoyar a los demás y fortalecer el avance colectivo.',
    'Desarrollo de relaciones':
      'Enfoque orientado a crear y sostener vínculos positivos basados en cercanía, confianza y comunicación abierta.',
    Empatía:
      'Disposición a comprender las perspectivas y emociones de otras personas, ajustando la interacción para fortalecer la confianza y la colaboración.',
    'Gestión de conflictos':
      'Manera de abordar desacuerdos y tensiones con equilibrio y claridad, interpretando las causas del problema y facilitando soluciones.',
    'Trabajo en equipo':
      'Capacidad de laborar y cooperar con otros miembros de la organización por un mismo objetivo, manteniendo o mejorando el propio desempeño.',
    'Trabajo en equipo por objetivos':
      'Enfoque colaborativo centrado en alcanzar metas comunes, integrando esfuerzos y manteniendo una visión compartida.',
    'Actitud Inclusiva (Cultural Diversity)':
      'Capacidad de valorar y promover la diversidad en todas sus formas, fomentando un entorno inclusivo, equitativo y respetuoso.',
    'Construcción de Relaciones de Confianza':
      'Capacidad de generar vínculos basados en credibilidad, honestidad e integridad, fomentando un ambiente de colaboración, respeto y transparencia.',
    'Sinergia Organizacional':
      'Capacidad de trabajar eficazmente con diversas áreas, integrando perspectivas y generando sinergias para alcanzar objetivos comunes.',
  },
  Liderazgo: {
    'Competencia del náufrago/sobrevivencia':
      'Permite mantener el funcionamiento efectivo en contextos de alta presión, incertidumbre o limitación de recursos.',
    'Creación de equipos de alto rendimiento':
      'Enfoque orientado a conformar equipos cohesionados y confiables, estableciendo condiciones que potencian la colaboración y el máximo desempeño.',
    'Desarrollo de equipo':
      'Orientación a fortalecer el desempeño colectivo mediante la creación de un entorno de colaboración, confianza y propósito común.',
    'Desarrollo de personas':
      'Favorece el crecimiento de otras personas mediante acompañamiento, retroalimentación constructiva y apertura para comprender distintas formas de pensar.',
    Dominancia:
      'Inclinación a asumir un rol protagónico en la toma de decisiones, expresarse con firmeza e influir en el rumbo de las acciones.',
    'Empowerment/Empoderamiento':
      'Orientación a impulsar la autonomía y el criterio propio de las personas, ofreciendo claridad y confianza para que tomen decisiones.',
    'Gestión de proyectos':
      'Capacidad para coordinar de manera integral tareas, recursos y tiempos, transformando objetivos en planes claros.',
    'Gestión del Riesgo':
      'Capacidad para analizar escenarios, identificar posibles amenazas y tomar decisiones informadas que protejan la operación y los recursos.',
    'Gestión del riesgo/seguridad':
      'Control riguroso de riesgos mediante decisiones preventivas alineadas a normas y procedimientos.',
    'Gestión estratégica del talento humano':
      'Enfoque integral para alinear las prácticas de talento con la estrategia organizacional, impulsando una cultura que favorece el desempeño.',
    'Gestión operativa':
      'Ejecución coordinada de procesos y tareas cotidianas, aplicando criterio para tomar decisiones oportunas y mantener control.',
    Intraemprendimiento:
      'Impulso para desarrollar ideas y proyectos dentro de la organización, explorando oportunidades y proponiendo mejoras con determinación.',
    'Liderazgo Afiliativo':
      'Estilo de liderazgo centrado en construir vínculos de confianza y cercanía, promoviendo bienestar emocional, armonía y un clima positivo.',
    'Liderazgo Coercitivo':
      'Estilo de liderazgo que ejerce control firme y establece reglas claras, exigiendo su cumplimiento para asegurar resultados inmediatos.',
    'Liderazgo Democrático':
      'Estilo de liderazgo que promueve la participación del equipo en la toma de decisiones, integrando ideas diversas.',
    'Liderazgo digital':
      'Liderazgo efectivo de equipos y operaciones en entornos virtuales, integrando criterio tecnológico y herramientas digitales.',
    'Liderazgo Facilitador':
      'Estilo de liderazgo que impulsa el desarrollo y la participación del equipo, removiendo obstáculos y creando condiciones para alcanzar resultados.',
    'Liderazgo Imitativo':
      'Lidera mediante el ejemplo, mostrando estándares altos de calidad y disciplina, y marcando ritmos exigentes.',
    'Liderazgo Orientativo':
      'Estilo de liderazgo que ofrece dirección clara y sentido de propósito, conectando las metas individuales con los objetivos colectivos.',
    'Liderazgo transformacional':
      'Estilo de liderazgo que inspira y moviliza al equipo hacia objetivos superiores, transmitiendo visión, propósito y energía positiva.',
    'Motivación Laboral':
      'Disposición a involucrarse con energía y compromiso en las tareas y metas del equipo, aportando una actitud positiva.',
    'Orientación al logro':
      'Impulso sostenido para alcanzar resultados y mantener altos estándares de desempeño, ajustando el esfuerzo hacia las metas.',
    'Potencial de liderazgo':
      'Inclinación a influir positivamente en las personas y orientar esfuerzos hacia objetivos comunes, mostrando claridad de rumbo.',
    'Supervisión y Control de Procesos':
      'Capacidad de supervisar el progreso de tareas y proyectos, asegurando el cumplimiento de plazos, objetivos y estándares de calidad.',
    'Delegación Efectiva':
      'Capacidad de asignar tareas estratégicamente, empoderando al equipo con orientación y recursos, fomentando la autonomía.',
    'Gestión eficiente de reuniones':
      'Capacidad de planificar, conducir y evaluar reuniones optimizando el tiempo, facilitando la participación y asegurando acuerdos concretos.',
    'Coaching efectivo':
      'Capacidad de guiar y potenciar el desarrollo de otros mediante escucha activa, retroalimentación y preguntas estratégicas.',
    'Liderazgo Transaccional':
      'Busca el cumplimiento de objetivos mediante normas bien definidas, supervisión y control. Se enfoca en resultados a corto plazo.',
    'Liderazgo Ágil (Agile Leadership)':
      'Promueve la adaptabilidad, la colaboración y la entrega rápida de valor. Empodera a equipos autoorganizados.',
    'Liderazgo Resiliente':
      'Moviliza a los equipos para enfrentar desafíos complejos sin soluciones claras, promoviendo aprendizaje colectivo e innovación.',
    'Liderazgo Situacional':
      'Ajusta el estilo del líder según la madurez y necesidades del equipo.',
    'Liderazgo Sinérgico':
      'Se basa en la colaboración profunda y el apoyo mutuo dentro del equipo, especialmente en situaciones de crisis o alta exigencia.',
    'Liderazgo Relacional':
      'Construye relaciones sólidas y de confianza como base para lograr los objetivos.',
    'Liderazgo Inclusivo':
      'Fomenta entornos donde todas las personas se sienten valoradas y pueden aportar desde su diversidad.',
    'Liderazgo de Alto Impacto':
      'Combina visión estratégica con ejecución consistente para generar resultados de alto valor.',
  },
  Valores: {
    'Compromiso y lealtad':
      'Implica la voluntad de alinear los propios intereses y comportamientos con las necesidades y objetivos de la empresa, desarrollando un sentido de pertenencia.',
    'Confiabilidad y congruencia':
      'Implica la coherencia entre nuestras acciones/comportamientos y nuestro discurso, teniendo en cuenta que ser congruente nos hace confiables frente a los demás.',
    'Consciencia Ambiental':
      'Compromiso con el cuidado del entorno y prácticas sostenibles en la gestión diaria del trabajo.',
    'Credibilidad técnica':
      'Reconocimiento basado en el dominio experto del conocimiento y la capacidad de sustentar decisiones con solidez técnica.',
    'Ética profesional':
      'Capacidad para actuar teniendo en cuenta los principios y ética de la profesión, llevándolos a la práctica en las actividades cotidianas.',
    Integridad:
      'Actuar con honestidad, coherencia y respeto por los principios, incluso en situaciones difíciles.',
    Justicia:
      'Capacidad de actuar con equidad e imparcialidad, tratando a las personas según lo que corresponde.',
    Responsabilidad:
      'Asumir compromisos y cumplirlos con diligencia, haciéndose cargo de las consecuencias de las propias acciones.',
    'Cumplimiento de Normas (Compliance)':
      'Capacidad de cumplir con normas, políticas y regulaciones de manera disciplinada y responsable, asegurando coherencia con los estándares organizativos y legales.',
    Accountability:
      'Rendición de cuentas sobre los resultados propios, con transparencia y disposición a asumir responsabilidades.',
  },
  Riesgos: {
    'Prevención de riesgos físicos':
      'Capacidad para identificar y prevenir riesgos físicos en el entorno laboral, aplicando medidas de seguridad adecuadas.',
    'Prevención de riesgos mecánicos':
      'Capacidad para anticipar y prevenir riesgos derivados del uso de maquinaria y equipos mecánicos.',
    'Prevención de riesgos de incendios y explosiones':
      'Capacidad para identificar y gestionar riesgos de incendio y explosión, aplicando protocolos preventivos.',
    'Prevención de riesgos en ciberseguridad':
      'Capacidad para identificar y prevenir amenazas digitales, aplicando buenas prácticas de ciberseguridad.',
    'Actitud Preventiva':
      'Disposición constante a anticipar y prevenir situaciones de riesgo en el entorno laboral.',
  },
};

// ============================================================
// HELPERS
// ============================================================
const getAllCompetencias = () => {
  const list = [];
  Object.entries(DICCIONARIO_EVALUAR).forEach(([categoria, comps]) => {
    Object.entries(comps).forEach(([nombre, definicion]) => {
      list.push({ categoria, nombre, definicion });
    });
  });
  return list;
};

const extractColorsFromImage = (file) => {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement('canvas');
      const w = Math.min(img.width, 150);
      const h = Math.min(img.height, 150);
      canvas.width = w;
      canvas.height = h;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0, w, h);
      try {
        const data = ctx.getImageData(0, 0, w, h).data;
        const colorMap = {};
        for (let i = 0; i < data.length; i += 16) {
          const r = Math.round(data[i] / 20) * 20;
          const g = Math.round(data[i + 1] / 20) * 20;
          const b = Math.round(data[i + 2] / 20) * 20;
          const a = data[i + 3];
          if (a < 200) continue;
          // Descartar blancos y negros puros
          const bright = (r + g + b) / 3;
          if (bright > 235 || bright < 25) continue;
          const key = `${r},${g},${b}`;
          colorMap[key] = (colorMap[key] || 0) + 1;
        }
        const sorted = Object.entries(colorMap)
          .sort((a, b) => b[1] - a[1])
          .slice(0, 5);
        const colors = sorted.map(([key]) => {
          const [r, g, b] = key.split(',').map(Number);
          return `#${[r, g, b]
            .map((x) => x.toString(16).padStart(2, '0'))
            .join('')}`;
        });
        resolve(colors.length > 0 ? colors : ['#2D3E2F', '#8B7355']);
      } catch (e) {
        resolve(['#2D3E2F', '#8B7355']);
      }
    };
    img.onerror = () => resolve(['#2D3E2F', '#8B7355']);
    const reader = new FileReader();
    reader.onload = (e) => {
      img.src = e.target.result;
    };
    reader.readAsDataURL(file);
  });
};

const fileToDataURL = (file) =>
  new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target.result);
    reader.readAsDataURL(file);
  });

// ============================================================
// PROMPT BUILDER
// ============================================================
// ============================================================
// EXTRACCIÓN DE TEXTO DE ARCHIVOS (PDF / DOCX / XLSX / CSV / TXT)
// ============================================================
const loadPdfJs = () =>
  new Promise((resolve, reject) => {
    if (window.pdfjsLib) return resolve(window.pdfjsLib);
    const script = document.createElement('script');
    script.src =
      'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
    script.onload = () => {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc =
        'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
      resolve(window.pdfjsLib);
    };
    script.onerror = () => reject(new Error('No se pudo cargar pdf.js'));
    document.head.appendChild(script);
  });

const extractTextFromPdf = async (file) => {
  const pdfjsLib = await loadPdfJs();
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  let out = '';
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const tc = await page.getTextContent();
    out += tc.items.map((it) => it.str).join(' ') + '\n\n';
  }
  return out.trim();
};

const extractTextFromDocx = async (file) => {
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.extractRawText({ arrayBuffer });
  return (result.value || '').trim();
};

const extractTextFromExcel = async (file) => {
  const arrayBuffer = await file.arrayBuffer();
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  let out = '';
  wb.SheetNames.forEach((name) => {
    const sheet = wb.Sheets[name];
    out += `\n=== Hoja: ${name} ===\n`;
    out += XLSX.utils.sheet_to_csv(sheet) + '\n';
  });
  return out.trim();
};

const extractTextFromTxt = async (file) => (await file.text()).trim();

// Router de extracción por extensión
const extractTextFromFile = async (file) => {
  const name = (file.name || '').toLowerCase();
  if (name.endsWith('.pdf')) return await extractTextFromPdf(file);
  if (name.endsWith('.docx')) return await extractTextFromDocx(file);
  if (
    name.endsWith('.xlsx') ||
    name.endsWith('.xls') ||
    name.endsWith('.csv')
  ) {
    return await extractTextFromExcel(file);
  }
  if (name.endsWith('.txt')) return await extractTextFromTxt(file);
  throw new Error(
    `Formato no soportado: ${file.name}. Usa PDF, Word (.docx), Excel (.xlsx/.csv) o texto (.txt).`
  );
};

// Formatea tamaño de archivo legible
const fmtSize = (bytes) => {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
};

const buildPrompt = (inputs) => {
  const dictCompact = Object.entries(DICCIONARIO_EVALUAR)
    .map(
      ([cat, comps]) =>
        `### ${cat}\n${Object.entries(comps)
          .map(([n, d]) => `- **${n}**: ${d}`)
          .join('\n')}`
    )
    .join('\n\n');

  const joinInsumo = (textVal, filesArr) => {
    const partes = [];
    if (textVal && textVal.trim()) partes.push(textVal.trim());
    if (filesArr && filesArr.length > 0) {
      filesArr.forEach((f) => {
        partes.push(
          `\n— Contenido extraído del archivo "${f.name}" —\n${f.extractedText}`
        );
      });
    }
    return partes.length > 0 ? partes.join('\n\n') : '(No provisto)';
  };

  const competenciasCargo = joinInsumo(
    inputs.competenciasCargo,
    inputs.files?.competenciasCargo
  );
  const competenciasCultura = joinInsumo(
    inputs.competenciasCultura,
    inputs.files?.competenciasCultura
  );
  const filosofiaOrganizacional = joinInsumo(
    inputs.filosofiaOrganizacional,
    inputs.files?.filosofiaOrganizacional
  );
  const infoAdicional = joinInsumo(
    inputs.infoAdicional,
    inputs.files?.infoAdicional
  );

  return `Eres experta senior en gestión de talento humano y homologación de competencias laborales para clientes LATAM. Tu tarea es homologar las competencias del cliente contra el diccionario Evaluar (Catálogo Marzo 2026).

## DICCIONARIO EVALUAR (fuente única de verdad — usa SOLO estas competencias)

${dictCompact}

## INFORMACIÓN DEL CLIENTE

**Cliente**: ${inputs.clientName}
**País**: ${inputs.country}
**Industria**: ${inputs.industria || 'No especificada'}

${
  inputs.industria
    ? `
## CONTEXTO DE MERCADO PARA LA INDUSTRIA

Como experta en ${inputs.industria} en ${inputs.country}, considera las competencias más valoradas y comunes en esta industria al momento de priorizar las homologaciones. Usa tu conocimiento del mercado laboral de ${inputs.country} para identificar qué competencias del diccionario Evaluar son más relevantes y frecuentemente solicitadas en ${inputs.industria}.

En tu respuesta, incluye una sección de comparativa de mercado que identifique:
- Las 5-7 competencias más comunes en ${inputs.industria} según el mercado de ${inputs.country}
- Cómo se comparan las competencias del cliente con las tendencias del mercado
- Brechas o fortalezas identificadas respecto al estándar de la industria
`
    : ''
}

## INSUMOS DEL CLIENTE

### 1) Competencias por cargo
${competenciasCargo}

### 2) Competencias de cultura
${competenciasCultura}

### 3) Filosofía organizacional (misión, visión, valores, cultura)
${filosofiaOrganizacional}

### 4) Información adicional relevante
${infoAdicional}

## REGLAS DE HOMOLOGACIÓN

### Estructura básica
1. Usa SOLO nombres de competencias que existan en el diccionario Evaluar. Copia el nombre EXACTO.
2. Categoriza cada competencia Evaluar usando su categoría del diccionario — esa categoría es el "Eje Evaluar".

### Homologación por niveles de cargo
3. **UNA competencia del cliente puede mapear a MÚLTIPLES competencias Evaluar según el nivel de cargo.**
   Ejemplo: "Pensamiento estratégico" del cliente puede homologarse así:
   - Nivel Ejecutor/Auxiliar → Conciencia organizacional (nivel desarrollo 5-6)
   - Nivel Mando Medio/Especialista → Pensamiento estratégico (nivel desarrollo 6-7)
   - Nivel Gerencial/Mando Alto → Visión de negocio (nivel desarrollo 7-8)
   
4. Niveles de cargo estándar:
   - **Ejecutor**: auxiliares, asistentes, analistas junior (nivel desarrollo: 5-6)
   - **Mando Medio**: especialistas, jefaturas primer nivel, coordinadores, analistas senior (nivel desarrollo: 6-7)
   - **Mando Alto**: gerencias, jefaturas segundo nivel (nivel desarrollo: 7-8)
   - **Directivo**: vicepresidencias, direcciones corporativas (nivel desarrollo: 8-9)

5. Nunca uses niveles menores a 5 (competencia no desarrollada).

### Priorización de competencias del diccionario (TODOS LOS EJES)
6. **REGLA GENERAL**: Prioriza SIEMPRE las competencias más comunes en el mercado y en la industria del cliente (${
    inputs.industria || 'general'
  }). Evita competencias rebuscadas o muy específicas salvo que el cliente las pida explícitamente.

7. **Liderazgo — Filtro estricto**: USA ÚNICAMENTE estos estilos:
   - Potencial de liderazgo
   - Liderazgo transformacional
   - Liderazgo de alto impacto
   - Liderazgo ágil
   - Liderazgo transaccional
   - Liderazgo facilitador
   - Liderazgo relacional
   **NO uses** otros estilos de liderazgo del diccionario (liderazgo situacional, liderazgo resonante, etc.) bajo ninguna circunstancia.

8. **Cognitivas — Priorizar estándar**: Pensamiento analítico, Inteligencia general, Resolución de problemas, Pensamiento crítico. Evita: Pensamiento abstracto, Pensamiento divergente.

9. **Todos los demás ejes** (Actitudinales, Comerciales, Interpersonales, etc.): Siempre prioriza las competencias más transversales y comúnmente usadas en la industria. Por ejemplo:
   - Orientación al cliente sobre Gestión de quejas
   - Trabajo en equipo sobre Integración de redes
   - Adaptabilidad sobre Resiliencia ante crisis
   Usa las más específicas solo si el cliente las pide o su contexto lo requiere claramente.

### Otras reglas
10. Si el cliente NO provee ejes de agrupación, propón 4-5 ejes coherentes para su cultura.
11. Si la definición del cliente NO está disponible, deja definicionCliente como cadena vacía "".
12. La justificación técnica debe conectar EXPLÍCITAMENTE la definición del cliente con la Evaluar, citando conceptos en común. 2-3 oraciones.
13. Si detectas familias de cargo, inclúyelas en nivelesPorFamilia.
14. Considera el contexto cultural del país (${
    inputs.country
  }) al interpretar matices.

## FORMATO DE SALIDA

Responde ÚNICAMENTE con un JSON válido (sin markdown, sin backticks), con esta estructura EXACTA:

{
  "resumenEjecutivo": "Resumen de 2-3 oraciones.",
  "ejesOrigen": "cliente" | "propuesto_ia",
  "ejesPropuestos": ["Eje 1", "Eje 2", ...],
  "homologaciones": [
    {
      "ejeCliente": "Nombre del eje del cliente",
      "competenciaCliente": "Nombre competencia cliente",
      "fraseCliente": "Frase/slogan si existe, sino cadena vacía",
      "definicionCliente": "Definición textual del cliente (o cadena vacía si no disponible)",
      "homologacionesPorNivel": [
        {
          "nivelCargo": "Ejecutor" | "Mando Medio" | "Mando Alto" | "Directivo",
          "ejeEvaluar": "Categoría del diccionario Evaluar",
          "competenciaEvaluar": "Nombre EXACTO del diccionario",
          "definicionEvaluar": "Definición textual del diccionario",
          "nivelDesarrolloExigido": 6,
          "justificacion": "Justificación técnica 2-3 oraciones."
        }
      ]
    }
  ],
  "nivelesPorFamilia": [
    {
      "familia": "Ejecutores | Mandos Medios | Mandos Altos | Directivos",
      "rangoMinimo": 5,
      "rangoMaximo": 6,
      "cargosDetectados": ["Cargo 1", "Cargo 2"]
    }
  ],
  "comparativaMercado": {
    "industria": "${inputs.industria || 'General'}",
    "competenciasComunesIdentificadas": ["Competencia 1", "Competencia 2", "..."],
    "analisis": "Análisis de 2-3 párrafos comparando las competencias del cliente con las tendencias del mercado en ${
      inputs.industria
    } (${
    inputs.country
  }). Menciona brechas, fortalezas y oportunidades de mejora.",
    "recomendaciones": ["Recomendación 1", "Recomendación 2"]
  }
}

Sé riguroso: nombres exactos del diccionario (priorizando competencias comunes en ${
    inputs.industria || 'la industria'
  }), homologaciones diferenciadas por nivel de cargo cuando aplique, niveles coherentes con la jerarquía.${
    inputs.industria
      ? ' Usa tu conocimiento de mercado para priorizar competencias relevantes en ' +
        inputs.industria +
        '.'
      : ''
  }`;
};

// ============================================================
// COMPONENTE PRINCIPAL
// ============================================================
export default function App() {
  const [step, setStep] = useState('home');
  const [clientData, setClientData] = useState({
    clientName: '',
    country: 'Ecuador',
    industria: '',
    cxUser: '',
    competenciasCargo: '',
    competenciasCultura: '',
    filosofiaOrganizacional: '',
    infoAdicional: '',
  });
  // Archivos adjuntos por campo: cada uno es { name, size, extractedText, status }
  const [clientFiles, setClientFiles] = useState({
    competenciasCargo: [],
    competenciasCultura: [],
    filosofiaOrganizacional: [],
    infoAdicional: [],
  });
  const [logoFile, setLogoFile] = useState(null);
  const [logoDataUrl, setLogoDataUrl] = useState(null);
  const [logoColors, setLogoColors] = useState([]);
  const [result, setResult] = useState(null);
  const [processing, setProcessing] = useState(false);
  const [progressMsg, setProgressMsg] = useState('');
  const [error, setError] = useState(null);
  const [history, setHistory] = useState([]);
  const [savedCxUser, setSavedCxUser] = useState('');

  // Cargar historial + usuario guardado al iniciar
  useEffect(() => {
    (async () => {
      try {
        const keys = await window.storage.list('homolog:');
        if (keys && keys.keys) {
          const items = [];
          for (const k of keys.keys) {
            try {
              const v = await window.storage.get(k);
              if (v) items.push(JSON.parse(v.value));
            } catch (e) {}
          }
          items.sort((a, b) => new Date(b.fecha) - new Date(a.fecha));
          setHistory(items);
        }
        const user = await window.storage.get('cx_user');
        if (user) {
          setSavedCxUser(user.value);
          setClientData((prev) => ({ ...prev, cxUser: user.value }));
        }
      } catch (e) {
        console.warn('Storage init error:', e);
      }
    })();
  }, []);

  // Cargar fuente Manrope
  useEffect(() => {
    const link = document.createElement('link');
    link.href =
      'https://fonts.googleapis.com/css2?family=Manrope:wght@200;300;400;500;600;700;800&display=swap';
    link.rel = 'stylesheet';
    document.head.appendChild(link);
  }, []);

  const handleLogoUpload = async (file) => {
    if (!file) return;
    setLogoFile(file);
    const dataUrl = await fileToDataURL(file);
    setLogoDataUrl(dataUrl);
    const colors = await extractColorsFromImage(file);
    setLogoColors(colors);
  };

  const generateHomologacion = async () => {
    setProcessing(true);
    setError(null);
    setStep('processing');
    setProgressMsg('Analizando competencias del cliente...');

    try {
      const prompt = buildPrompt({ ...clientData, files: clientFiles });
      setProgressMsg(
        'Homologando contra diccionario Evaluar (149 competencias)...'
      );

      const response = await window.claude.complete(prompt);
      setProgressMsg('Estructurando resultado...');

      // Extraer JSON limpio
      let jsonText = response.trim();
      jsonText = jsonText
        .replace(/^```json\s*/i, '')
        .replace(/^```\s*/i, '')
        .replace(/\s*```\s*$/i, '');
      const firstBrace = jsonText.indexOf('{');
      const lastBrace = jsonText.lastIndexOf('}');
      if (firstBrace !== -1 && lastBrace !== -1) {
        jsonText = jsonText.substring(firstBrace, lastBrace + 1);
      }

      const parsed = JSON.parse(jsonText);
      const now = new Date();
      const finalResult = {
        ...parsed,
        id: `${now.getTime()}`,
        cliente: clientData.clientName,
        pais: clientData.country,
        industria: clientData.industria,
        cxUser: clientData.cxUser,
        fecha: now.toISOString(),
        logoDataUrl,
        logoColors,
      };

      setResult(finalResult);
      setProcessing(false);
      setStep('preview');

      // Guardar usuario CX
      if (clientData.cxUser) {
        await window.storage.set('cx_user', clientData.cxUser);
      }
    } catch (err) {
      console.error(err);
      setError(
        `Error al generar la homologación: ${err.message}. Revisa que los datos estén completos y vuelve a intentar.`
      );
      setProcessing(false);
      setStep('new');
    }
  };

  const saveToRepository = async () => {
    if (!result) return;
    try {
      const key = `homolog:${result.id}`;
      await window.storage.set(key, JSON.stringify(result));
      const keys = await window.storage.list('homolog:');
      if (keys && keys.keys) {
        const items = [];
        for (const k of keys.keys) {
          try {
            const v = await window.storage.get(k);
            if (v) items.push(JSON.parse(v.value));
          } catch (e) {}
        }
        items.sort((a, b) => new Date(b.fecha) - new Date(a.fecha));
        setHistory(items);
      }
      return true;
    } catch (e) {
      console.error('Save error:', e);
      return false;
    }
  };

  const exportToExcel = async () => {
    if (!result) return;
    await saveToRepository();

    const wb = XLSX.utils.book_new();

    // --- Pestaña 1: Portada e Insumos ---
    const portadaData = [
      ['HOMOLOGACIÓN DE COMPETENCIAS'],
      [],
      ['Cliente', result.cliente],
      ['País', result.pais || '—'],
      ['Industria', result.industria || '—'],
      [
        'Fecha de homologación',
        new Date(result.fecha).toLocaleDateString('es-EC', {
          year: 'numeric',
          month: 'long',
          day: 'numeric',
        }),
      ],
      ['Responsable CX', result.cxUser || '—'],
      [
        'Origen de ejes',
        result.ejesOrigen === 'cliente'
          ? 'Provistos por el cliente'
          : 'Propuestos por IA',
      ],
      [],
      ['RESUMEN EJECUTIVO'],
      [result.resumenEjecutivo],
      [],
      ['EJES DE COMPETENCIAS'],
      ...(result.ejesPropuestos || []).map((e, i) => [`${i + 1}.`, e]),
      [],
      ['COLORES DE MARCA EXTRAÍDOS DEL LOGO'],
      ...(result.logoColors || []).map((c, i) => [`Color ${i + 1}`, c]),
    ];
    const ws1 = XLSX.utils.aoa_to_sheet(portadaData);
    ws1['!cols'] = [{ wch: 30 }, { wch: 80 }];
    XLSX.utils.book_append_sheet(wb, ws1, 'Portada e Insumos');

    // --- Pestaña 2: Matriz de Homologación (9 columnas) ---
    const matrizHeaders = [
      'Eje cliente',
      'Competencia cliente',
      'Definición cliente',
      'Nivel de cargo',
      'Eje Evaluar',
      'Competencia Evaluar',
      'Definición Evaluar',
      'Nivel de desarrollo a exigir',
      'Justificación',
    ];
    const matrizRows = [matrizHeaders];
    (result.homologaciones || []).forEach((h) => {
      // Soporte para nueva estructura (homologacionesPorNivel) y legacy (homologacionesEvaluar)
      const evaluarList =
        h.homologacionesPorNivel || h.homologacionesEvaluar || [];
      evaluarList.forEach((e, idx) => {
        const ejeEv = e.ejeEvaluar || e.categoriaEvaluar || '';
        const nivel = e.nivelDesarrolloExigido ?? e.nivelDesarrollo ?? '';
        const nivelCargo = e.nivelCargo || h.nivelCargoCliente || '';
        matrizRows.push([
          idx === 0 ? h.ejeCliente : '',
          idx === 0 ? h.competenciaCliente : '',
          idx === 0 ? h.definicionCliente || '' : '',
          nivelCargo,
          ejeEv,
          e.competenciaEvaluar,
          e.definicionEvaluar,
          nivel,
          e.justificacion,
        ]);
      });
    });
    const ws2 = XLSX.utils.aoa_to_sheet(matrizRows);
    ws2['!cols'] = [
      { wch: 22 },
      { wch: 28 },
      { wch: 50 },
      { wch: 18 },
      { wch: 22 },
      { wch: 28 },
      { wch: 50 },
      { wch: 18 },
      { wch: 55 },
    ];
    XLSX.utils.book_append_sheet(wb, ws2, 'Matriz de Homologación');

    // --- Pestaña 3: Niveles por familia ---
    if (result.nivelesPorFamilia && result.nivelesPorFamilia.length > 0) {
      const nivelesData = [
        [
          'Familia de cargo',
          'Nivel mínimo',
          'Nivel máximo',
          'Cargos detectados',
        ],
        ...result.nivelesPorFamilia.map((n) => [
          n.familia,
          n.rangoMinimo,
          n.rangoMaximo,
          (n.cargosDetectados || []).join(' · '),
        ]),
      ];
      const ws3 = XLSX.utils.aoa_to_sheet(nivelesData);
      ws3['!cols'] = [{ wch: 25 }, { wch: 15 }, { wch: 15 }, { wch: 60 }];
      XLSX.utils.book_append_sheet(wb, ws3, 'Niveles por familia');
    }

    const safeName = result.cliente.replace(/[^a-z0-9]/gi, '_');
    const dateStr = new Date(result.fecha).toISOString().slice(0, 10);
    XLSX.writeFile(wb, `Homologacion_${safeName}_${dateStr}.xlsx`);
  };

  const exportToPDF = async () => {
    if (!result) return;
    await saveToRepository();

    const doc = new jsPDF('l', 'pt', 'a4'); // Landscape, points, A4
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = 40;

    // Header con logo del cliente si existe
    if (result.logoDataUrl) {
      try {
        doc.addImage(result.logoDataUrl, 'PNG', margin, margin, 80, 40);
      } catch (e) {
        console.warn('No se pudo agregar logo al PDF:', e);
      }
    }

    // Título y metadatos
    doc.setFontSize(20);
    doc.setTextColor(46, 28, 90); // Brand color
    doc.text('HOMOLOGACIÓN DE COMPETENCIAS', pageWidth / 2, margin + 20, {
      align: 'center',
    });

    doc.setFontSize(10);
    doc.setTextColor(107, 100, 128); // Muted
    doc.text(`Cliente: ${result.cliente}`, margin, margin + 60);
    doc.text(`País: ${result.pais || '—'}`, margin, margin + 75);
    if (result.industria)
      doc.text(`Industria: ${result.industria}`, margin, margin + 90);
    doc.text(
      `Fecha: ${new Date(result.fecha).toLocaleDateString('es-EC')}`,
      pageWidth - margin,
      margin + 60,
      { align: 'right' }
    );
    doc.text(`CX: ${result.cxUser || '—'}`, pageWidth - margin, margin + 75, {
      align: 'right',
    });

    // Resumen ejecutivo
    let y = margin + 120;
    doc.setFontSize(12);
    doc.setTextColor(46, 28, 90);
    doc.text('Resumen ejecutivo:', margin, y);
    doc.setFontSize(9);
    doc.setTextColor(60, 60, 60);
    const resumenLines = doc.splitTextToSize(
      result.resumenEjecutivo || '',
      pageWidth - 2 * margin
    );
    doc.text(resumenLines, margin, y + 15);
    y += 15 + resumenLines.length * 12 + 25;

    // Tabla de homologación con autoTable
    const tableHeaders = [
      'Eje cliente',
      'Competencia cliente',
      'Def. cliente',
      'Nivel cargo',
      'Eje Evaluar',
      'Competencia Evaluar',
      'Def. Evaluar',
      'Nivel',
      'Justificación',
    ];
    const tableRows = [];
    (result.homologaciones || []).forEach((h) => {
      const evaluarList =
        h.homologacionesPorNivel || h.homologacionesEvaluar || [];
      evaluarList.forEach((e, idx) => {
        const ejeEv = e.ejeEvaluar || e.categoriaEvaluar || '';
        const nivel = e.nivelDesarrolloExigido ?? e.nivelDesarrollo ?? '';
        const nivelCargo = e.nivelCargo || h.nivelCargoCliente || '';
        tableRows.push([
          idx === 0 ? h.ejeCliente : '',
          idx === 0 ? h.competenciaCliente : '',
          idx === 0 ? h.definicionCliente || '—' : '',
          nivelCargo || '—',
          ejeEv,
          e.competenciaEvaluar,
          e.definicionEvaluar,
          nivel,
          e.justificacion,
        ]);
      });
    });

    doc.autoTable({
      startY: y,
      head: [tableHeaders],
      body: tableRows,
      theme: 'grid',
      styles: { fontSize: 7, cellPadding: 4 },
      headStyles: { fillColor: [46, 28, 90], textColor: 255 },
      columnStyles: {
        0: { cellWidth: 50 }, // Eje cliente
        1: { cellWidth: 70 }, // Competencia cliente
        2: { cellWidth: 90 }, // Def cliente
        3: { cellWidth: 50 }, // Nivel cargo
        4: { cellWidth: 60 }, // Eje Evaluar
        5: { cellWidth: 80 }, // Competencia Evaluar
        6: { cellWidth: 100 }, // Def Evaluar
        7: { cellWidth: 30 }, // Nivel
        8: { cellWidth: 120 }, // Justificación
      },
      margin: { left: margin, right: margin },
    });

    // Footer con logo Evaluar
    const totalPages = doc.internal.getNumberOfPages();
    for (let i = 1; i <= totalPages; i++) {
      doc.setPage(i);
      doc.setFontSize(8);
      doc.setTextColor(150, 150, 150);
      doc.text(`Página ${i} de ${totalPages}`, pageWidth / 2, pageHeight - 20, {
        align: 'center',
      });
      doc.text('Generado por evaluar.ai', pageWidth - margin, pageHeight - 20, {
        align: 'right',
      });
    }

    const safeName = result.cliente.replace(/[^a-z0-9]/gi, '_');
    const dateStr = new Date(result.fecha).toISOString().slice(0, 10);
    doc.save(`Homologacion_${safeName}_${dateStr}.pdf`);
  };

  const resetFlow = () => {
    setClientData({
      clientName: '',
      country: 'Ecuador',
      industria: '',
      cxUser: savedCxUser,
      competenciasCargo: '',
      competenciasCultura: '',
      filosofiaOrganizacional: '',
      infoAdicional: '',
    });
    setClientFiles({
      competenciasCargo: [],
      competenciasCultura: [],
      filosofiaOrganizacional: [],
      infoAdicional: [],
    });
    setLogoFile(null);
    setLogoDataUrl(null);
    setLogoColors([]);
    setResult(null);
    setError(null);
    setStep('home');
  };

  // ============================================================
  // RENDER
  // ============================================================
  return (
    <div
      className="min-h-screen noise-bg"
      style={{
        backgroundColor: '#FAF8FC',
        fontFamily: "'Manrope', sans-serif",
      }}
    >
      <style>{`
        * { font-family: 'Manrope', sans-serif; }
        .font-display { font-family: 'Manrope', sans-serif; font-weight: 800; letter-spacing: -0.04em; }
        .font-accent { font-family: 'Manrope', sans-serif; font-weight: 300; font-style: italic; letter-spacing: -0.02em; }
        .font-body { font-family: 'Manrope', sans-serif; font-weight: 400; }
        .text-ink { color: #1A1033; }
        .text-muted { color: #6B6480; }
        .text-brand { color: #2E1C5A; }
        .text-magenta { color: #D91B7E; }
        .text-coral { color: #F5A623; }
        .bg-canvas { background-color: #FAF8FC; }
        .bg-paper { background-color: #FFFFFF; }
        .bg-ink { background-color: #1A1033; }
        .bg-brand { background-color: #2E1C5A; }
        .bg-magenta { background-color: #D91B7E; }
        .bg-coral { background-color: #F5A623; }
        .bg-tint { background-color: #F3EFFB; }
        .border-ink { border-color: #1A1033; }
        .border-soft { border-color: #E6E1F0; }
        .border-brand { border-color: #2E1C5A; }
        .hover-shift { transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1); }
        .hover-shift:hover { transform: translateY(-2px); }
        .gradient-brand { background: linear-gradient(135deg, #2E1C5A 0%, #5B2B9C 100%); }
        .gradient-accent { background: linear-gradient(135deg, #D91B7E 0%, #F5A623 100%); }
        .gradient-text { background: linear-gradient(135deg, #2E1C5A 0%, #D91B7E 60%, #F5A623 100%); -webkit-background-clip: text; background-clip: text; color: transparent; }
        .shadow-brand { box-shadow: 0 1px 3px rgba(46,28,90,0.04), 0 4px 20px rgba(46,28,90,0.08); }
        .shadow-brand-lg { box-shadow: 0 4px 24px rgba(46,28,90,0.08), 0 12px 48px rgba(46,28,90,0.12); }
        textarea, input { font-family: 'Manrope', sans-serif; }
        textarea:focus, input:focus { outline: none; border-color: #2E1C5A; box-shadow: 0 0 0 3px rgba(46,28,90,0.08); }
        .noise-bg { background-image: radial-gradient(circle at 20% 10%, rgba(217,27,126,0.04) 0%, transparent 50%), radial-gradient(circle at 80% 80%, rgba(245,166,35,0.05) 0%, transparent 50%); }
      `}</style>

      {/* HEADER */}
      <header
        className="border-b border-soft bg-paper sticky top-0 z-50 backdrop-blur"
        style={{ backgroundColor: 'rgba(255,255,255,0.92)' }}
      >
        <div className="max-w-7xl mx-auto px-8 py-4 flex items-center justify-between">
          <div className="flex items-center gap-5">
            <EvaluarLogo className="h-9" />
            <div className="h-8 w-px bg-soft" />
            <div>
              <div
                className="text-lg text-ink leading-none"
                style={{ fontWeight: 700, letterSpacing: '-0.025em' }}
              >
                Homologador<span className="text-magenta">.</span>
              </div>
              <div className="text-[11px] text-muted mt-1 tracking-[0.15em] uppercase">
                Customer Experience
              </div>
            </div>
          </div>
          <nav className="flex items-center gap-1">
            <button
              onClick={() => setStep('home')}
              className={`px-4 py-2 text-sm rounded-full hover-shift ${
                step === 'home'
                  ? 'bg-brand text-white'
                  : 'text-ink hover:bg-tint'
              }`}
            >
              Inicio
            </button>
            <button
              onClick={() => setStep('history')}
              className={`px-4 py-2 text-sm rounded-full hover-shift flex items-center gap-2 ${
                step === 'history'
                  ? 'bg-brand text-white'
                  : 'text-ink hover:bg-tint'
              }`}
            >
              <Archive size={14} /> Repositorio{' '}
              <span
                className={`${
                  step === 'history' ? 'text-white' : 'text-magenta'
                } font-semibold`}
              >
                ({history.length})
              </span>
            </button>
          </nav>
        </div>
      </header>

      {/* MAIN */}
      <main className="max-w-7xl mx-auto px-8 py-12">
        {step === 'home' && <HomeView setStep={setStep} history={history} />}
        {step === 'new' && (
          <NewHomologationView
            clientData={clientData}
            setClientData={setClientData}
            clientFiles={clientFiles}
            setClientFiles={setClientFiles}
            logoDataUrl={logoDataUrl}
            logoColors={logoColors}
            handleLogoUpload={handleLogoUpload}
            generateHomologacion={generateHomologacion}
            error={error}
            setStep={setStep}
          />
        )}
        {step === 'processing' && <ProcessingView progressMsg={progressMsg} />}
        {step === 'preview' && result && (
          <PreviewView
            result={result}
            setResult={setResult}
            exportToExcel={exportToExcel}
            resetFlow={resetFlow}
          />
        )}
        {step === 'history' && (
          <HistoryView
            history={history}
            setResult={setResult}
            setStep={setStep}
          />
        )}
      </main>

      <footer className="border-t border-soft mt-24 py-10">
        <div className="max-w-7xl mx-auto px-8 flex items-center justify-between flex-wrap gap-4">
          <div className="flex items-center gap-3">
            <EvaluarLogo className="h-6" opacity={0.6} />
            <span className="text-xs text-muted">· Homologador CX</span>
          </div>
          <p className="text-[11px] text-muted tracking-wide">
            Prototipo de validación · Lógica base para la Web App en Next.js
          </p>
        </div>
      </footer>
    </div>
  );
}

// ============================================================
// HOME VIEW
// ============================================================
function HomeView({ setStep, history }) {
  const stats = {
    total: history.length,
    uniqueClients: new Set(history.map((h) => h.cliente)).size,
    thisMonth: history.filter(
      (h) => new Date(h.fecha).getMonth() === new Date().getMonth()
    ).length,
  };

  return (
    <div className="space-y-14">
      {/* Hero */}
      <div className="grid grid-cols-12 gap-8 items-end pt-10">
        <div className="col-span-12 md:col-span-8">
          <div className="inline-flex items-center gap-2 px-3 py-1.5 bg-tint rounded-full mb-6">
            <div className="w-1.5 h-1.5 rounded-full bg-magenta animate-pulse" />
            <span className="text-xs text-brand tracking-widest uppercase font-semibold">
              Gestión de talento · LATAM
            </span>
          </div>
          <h1
            className="font-display text-ink"
            style={{ fontSize: 'clamp(2.5rem, 5vw, 4.75rem)', lineHeight: '1' }}
          >
            Homologación de
            <br />
            <span className="gradient-text">competencias</span>,<br />
            hecha con criterio.
          </h1>
          <p className="mt-8 text-lg text-muted max-w-2xl leading-relaxed">
            Traduce el lenguaje de competencias de tu cliente al diccionario
            Evaluar. Con justificación técnica, niveles coherentes y entregable
            listo para Drive.
          </p>
        </div>
        <div className="col-span-12 md:col-span-4 bg-paper border border-soft rounded-2xl p-7 shadow-brand">
          <div className="flex items-baseline justify-between mb-1">
            <div>
              <div className="text-[11px] text-muted tracking-[0.18em] uppercase mb-1">
                Diccionario base
              </div>
              <div className="font-display text-5xl text-brand">149</div>
              <div className="text-sm text-muted mt-1">
                competencias Evaluar
              </div>
            </div>
          </div>
          <div className="h-px bg-soft my-5" />
          <div className="flex items-baseline justify-between">
            <div>
              <div className="text-[11px] text-muted tracking-[0.18em] uppercase mb-1">
                Categorías
              </div>
              <div
                className="font-display text-5xl"
                style={{ color: '#D91B7E' }}
              >
                12
              </div>
              <div className="text-sm text-muted mt-1">ejes temáticos</div>
            </div>
          </div>
        </div>
      </div>

      {/* Tarjetas de acción */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
        <button
          onClick={() => setStep('new')}
          className="gradient-brand text-white p-8 rounded-2xl text-left hover-shift group shadow-brand-lg relative overflow-hidden"
        >
          <div
            className="absolute top-0 right-0 w-32 h-32 rounded-full opacity-20"
            style={{
              background:
                'radial-gradient(circle, #F5A623 0%, transparent 70%)',
              transform: 'translate(40%, -40%)',
            }}
          />
          <div
            className="absolute bottom-0 left-0 w-24 h-24 rounded-full opacity-15"
            style={{
              background:
                'radial-gradient(circle, #D91B7E 0%, transparent 70%)',
              transform: 'translate(-30%, 30%)',
            }}
          />
          <div className="relative">
            <div className="flex items-center justify-between mb-8">
              <div className="w-11 h-11 rounded-xl bg-white bg-opacity-15 flex items-center justify-center backdrop-blur">
                <Sparkles size={20} strokeWidth={1.8} />
              </div>
              <ChevronRight
                size={20}
                className="opacity-70 group-hover:translate-x-1 transition-transform"
              />
            </div>
            <div
              className="font-display text-3xl mb-2"
              style={{ letterSpacing: '-0.03em' }}
            >
              Nueva homologación
            </div>
            <div className="text-sm opacity-85 leading-relaxed">
              Carga insumos del cliente y genera la homologación completa con
              IA.
            </div>
          </div>
        </button>

        <button
          onClick={() => setStep('history')}
          className="bg-paper border border-soft p-8 rounded-2xl text-left hover-shift group shadow-brand"
        >
          <div className="flex items-center justify-between mb-8">
            <div className="w-11 h-11 rounded-xl bg-tint flex items-center justify-center">
              <Archive size={20} strokeWidth={1.8} className="text-brand" />
            </div>
            <ChevronRight
              size={20}
              className="text-muted group-hover:translate-x-1 transition-transform"
            />
          </div>
          <div
            className="font-display text-3xl text-ink mb-2"
            style={{ letterSpacing: '-0.03em' }}
          >
            Repositorio
          </div>
          <div className="text-sm text-muted leading-relaxed">
            <span className="text-brand font-semibold">{stats.total}</span>{' '}
            homologaciones ·{' '}
            <span className="text-brand font-semibold">
              {stats.uniqueClients}
            </span>{' '}
            clientes ·{' '}
            <span className="text-brand font-semibold">{stats.thisMonth}</span>{' '}
            este mes
          </div>
        </button>

        <div className="bg-paper border border-soft p-8 rounded-2xl shadow-brand">
          <div className="flex items-center justify-between mb-6">
            <div className="text-[11px] text-muted tracking-[0.18em] uppercase font-semibold">
              Flujo
            </div>
            <div className="w-11 h-11 rounded-xl bg-tint flex items-center justify-center">
              <FileText size={18} className="text-brand" strokeWidth={1.8} />
            </div>
          </div>
          <ol className="space-y-3.5 text-sm text-ink">
            <li className="flex gap-3 items-start">
              <span className="font-display text-brand w-7">01</span> Carga
              insumos del cliente
            </li>
            <li className="flex gap-3 items-start">
              <span className="font-display text-brand w-7">02</span> IA
              homologa contra Evaluar
            </li>
            <li className="flex gap-3 items-start">
              <span className="font-display text-brand w-7">03</span>{' '}
              Previsualiza y edita
            </li>
            <li className="flex gap-3 items-start">
              <span className="font-display text-brand w-7">04</span> Descarga
              Excel + repositorio
            </li>
          </ol>
        </div>
      </div>
    </div>
  );
}

// ============================================================
// NEW HOMOLOGATION VIEW
// ============================================================
const PAISES_LATAM = ['Ecuador', 'Colombia', 'Chile', 'Perú', 'México'];

const CAMPOS_INSUMOS = [
  {
    key: 'competenciasCargo',
    titulo: 'Competencias por cargo',
    descripcion:
      'Competencias específicas asociadas a los cargos o grupos ocupacionales del cliente.',
    placeholder: `Ejemplo:

Director Comercial:
- Visión de negocio: lee tendencias del mercado...
- Negociación estratégica: construye acuerdos de largo plazo...

Analista de Ventas:
- Análisis de datos comerciales...`,
  },
  {
    key: 'competenciasCultura',
    titulo: 'Competencias de cultura',
    descripcion:
      'Competencias transversales que reflejan la cultura del cliente (suelen aplicar a toda la organización).',
    placeholder: `Ejemplo:

- Pasión por el cliente: el cliente es el centro de todo lo que hacemos...
- Excelencia integral: hacemos las cosas bien, siempre...
- Colaboración: trabajamos como un solo equipo...`,
  },
  {
    key: 'filosofiaOrganizacional',
    titulo: 'Filosofía organizacional',
    descripcion:
      'Misión, visión, valores, propósito y cultura organizacional del cliente.',
    placeholder: `Ejemplo:

Misión: ...
Visión: ...
Valores: ...
Propósito: ...`,
  },
  {
    key: 'infoAdicional',
    titulo: 'Información adicional relevante',
    descripcion:
      'Organigrama, estrategia, contexto de industria, glosarios, o cualquier otro insumo útil.',
    placeholder: `Ejemplo:

Organigrama: Director Nacional → Gerentes regionales → Jefes de área → Analistas...

Contexto de industria: empresa del sector retail con 5000 colaboradores...`,
  },
];

function FileChip({ file, onRemove }) {
  const isError = file.status === 'error';
  const isLoading = file.status === 'loading';
  return (
    <div
      className={`flex items-center gap-2 px-3 py-2 rounded-lg border text-xs ${
        isError ? 'bg-red-50 border-red-200' : 'bg-tint border-soft'
      }`}
    >
      {isLoading ? (
        <Loader2 size={14} className="animate-spin text-brand flex-shrink-0" />
      ) : (
        <FileIcon
          size={14}
          className={`flex-shrink-0 ${isError ? 'text-red-600' : 'text-brand'}`}
        />
      )}
      <div className="min-w-0 flex-1">
        <div
          className={`truncate ${
            isError ? 'text-red-700' : 'text-ink'
          } font-medium`}
        >
          {file.name}
        </div>
        <div
          className={`truncate text-[10px] ${
            isError ? 'text-red-600' : 'text-muted'
          }`}
        >
          {isError
            ? file.error
            : isLoading
            ? 'Procesando...'
            : `${fmtSize(file.size)} · ${
                file.extractedText?.length || 0
              } caracteres extraídos`}
        </div>
      </div>
      {!isLoading && (
        <button
          onClick={onRemove}
          className="text-muted hover:text-red-600 flex-shrink-0"
          type="button"
        >
          <X size={14} />
        </button>
      )}
    </div>
  );
}

function InsumoField({
  fieldConfig,
  value,
  files,
  onTextChange,
  onFilesAdd,
  onFileRemove,
}) {
  const fileInputRef = useRef(null);
  const [dragOver, setDragOver] = useState(false);

  const handleFilesPicked = async (fileList) => {
    const arr = Array.from(fileList || []);
    if (arr.length === 0) return;
    await onFilesAdd(arr);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setDragOver(false);
    handleFilesPicked(e.dataTransfer.files);
  };

  return (
    <section className="bg-paper border border-soft rounded-2xl p-8 shadow-brand">
      <div className="flex items-start justify-between gap-4 mb-2">
        <div className="flex-1">
          <h3
            className="text-xl text-ink"
            style={{ fontWeight: 700, letterSpacing: '-0.02em' }}
          >
            {fieldConfig.titulo}
          </h3>
          <p className="text-sm text-muted mt-1 leading-relaxed">
            {fieldConfig.descripcion}
          </p>
        </div>
        <span className="text-[10px] text-muted tracking-[0.18em] uppercase font-semibold bg-tint px-2.5 py-1 rounded-full whitespace-nowrap">
          Opcional
        </span>
      </div>

      <textarea
        value={value}
        onChange={(e) => onTextChange(e.target.value)}
        placeholder={fieldConfig.placeholder}
        rows={6}
        className="w-full px-4 py-3 bg-canvas border border-soft rounded-xl text-ink text-sm leading-relaxed transition-all mt-4"
      />

      {/* Zona de archivos */}
      <div
        onDragOver={(e) => {
          e.preventDefault();
          setDragOver(true);
        }}
        onDragLeave={() => setDragOver(false)}
        onDrop={handleDrop}
        className={`mt-3 border-2 border-dashed rounded-xl p-4 transition-all ${
          dragOver ? 'border-brand bg-tint' : 'border-soft bg-canvas'
        }`}
      >
        <div className="flex items-center justify-between gap-3 flex-wrap">
          <div className="flex items-center gap-2 text-xs text-muted">
            <Paperclip size={14} className="text-brand" />
            <span>Arrastra archivos o</span>
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              className="text-brand hover:text-magenta font-semibold underline underline-offset-2"
            >
              haz clic para subir
            </button>
          </div>
          <span className="text-[10px] text-muted">
            PDF · Word · Excel · CSV · TXT
          </span>
        </div>
        <input
          ref={fileInputRef}
          type="file"
          multiple
          accept=".pdf,.docx,.xlsx,.xls,.csv,.txt"
          onChange={(e) => handleFilesPicked(e.target.files)}
          className="hidden"
        />
        {files.length > 0 && (
          <div className="grid grid-cols-1 md:grid-cols-2 gap-2 mt-3">
            {files.map((f, i) => (
              <FileChip key={i} file={f} onRemove={() => onFileRemove(i)} />
            ))}
          </div>
        )}
      </div>
    </section>
  );
}

function NewHomologationView({
  clientData,
  setClientData,
  clientFiles,
  setClientFiles,
  logoDataUrl,
  logoColors,
  handleLogoUpload,
  generateHomologacion,
  error,
  setStep,
}) {
  const fileInputRef = useRef(null);

  // ¿Al menos uno de los 4 campos tiene contenido (texto o archivo)?
  const hasContent = (key) => {
    const hasText = clientData[key] && clientData[key].trim().length > 0;
    const hasValidFile = (clientFiles[key] || []).some(
      (f) => f.status === 'ready'
    );
    return hasText || hasValidFile;
  };
  const atLeastOneInsumo = [
    'competenciasCargo',
    'competenciasCultura',
    'filosofiaOrganizacional',
    'infoAdicional',
  ].some(hasContent);

  const canSubmit =
    clientData.clientName.trim() &&
    clientData.country &&
    clientData.cxUser.trim() &&
    atLeastOneInsumo;

  // Handlers de archivos por campo
  const addFilesToField = async (fieldKey, fileArr) => {
    // Primero los añadimos como "loading"
    const placeholders = fileArr.map((f) => ({
      name: f.name,
      size: f.size,
      status: 'loading',
      extractedText: '',
    }));
    setClientFiles((prev) => ({
      ...prev,
      [fieldKey]: [...prev[fieldKey], ...placeholders],
    }));

    // Procesamos uno a uno (para no saturar memoria con PDFs grandes)
    for (let i = 0; i < fileArr.length; i++) {
      const file = fileArr[i];
      try {
        const extractedText = await extractTextFromFile(file);
        setClientFiles((prev) => {
          const arr = [...prev[fieldKey]];
          // Encontrar el placeholder que corresponde (por nombre+tamaño, marcado como loading)
          const idx = arr.findIndex(
            (x) =>
              x.name === file.name &&
              x.size === file.size &&
              x.status === 'loading'
          );
          if (idx !== -1) {
            arr[idx] = {
              name: file.name,
              size: file.size,
              status: 'ready',
              extractedText,
            };
          }
          return { ...prev, [fieldKey]: arr };
        });
      } catch (err) {
        setClientFiles((prev) => {
          const arr = [...prev[fieldKey]];
          const idx = arr.findIndex(
            (x) =>
              x.name === file.name &&
              x.size === file.size &&
              x.status === 'loading'
          );
          if (idx !== -1) {
            arr[idx] = {
              name: file.name,
              size: file.size,
              status: 'error',
              extractedText: '',
              error: err.message,
            };
          }
          return { ...prev, [fieldKey]: arr };
        });
      }
    }
  };

  const removeFileFromField = (fieldKey, index) => {
    setClientFiles((prev) => ({
      ...prev,
      [fieldKey]: prev[fieldKey].filter((_, i) => i !== index),
    }));
  };

  return (
    <div className="max-w-5xl mx-auto">
      <button
        onClick={() => setStep('home')}
        className="text-sm text-muted flex items-center gap-2 mb-8 hover:text-brand hover-shift"
      >
        <ChevronLeft size={16} /> Volver al inicio
      </button>

      <div className="mb-12">
        <div className="inline-flex items-center gap-2 px-3 py-1 bg-tint rounded-full mb-4">
          <span className="text-[11px] text-brand tracking-[0.18em] uppercase font-semibold">
            Paso 01 · Insumos del cliente
          </span>
        </div>
        <h2
          className="font-display text-ink"
          style={{ fontSize: 'clamp(2rem, 4vw, 3.5rem)', lineHeight: '1.05' }}
        >
          Describamos al cliente.
        </h2>
      </div>

      {error && (
        <div className="bg-red-50 border border-red-200 rounded-sm p-4 mb-8 flex gap-3">
          <AlertCircle
            size={18}
            className="text-red-600 flex-shrink-0 mt-0.5"
          />
          <div className="text-sm text-red-700">{error}</div>
        </div>
      )}

      <div className="space-y-8">
        {/* Datos básicos */}
        <section className="bg-paper border border-soft rounded-2xl p-8 shadow-brand">
          <div className="flex items-center gap-2 mb-6">
            <Building2 size={18} className="text-brand" />
            <h3
              className="text-xl text-ink"
              style={{ fontWeight: 700, letterSpacing: '-0.02em' }}
            >
              Datos básicos
            </h3>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div>
              <label className="block text-[11px] text-muted tracking-[0.18em] uppercase mb-2 font-semibold">
                Nombre del cliente <span className="text-magenta">*</span>
              </label>
              <input
                type="text"
                value={clientData.clientName}
                onChange={(e) =>
                  setClientData({ ...clientData, clientName: e.target.value })
                }
                placeholder="Ej: Corporación Favorita"
                className="w-full px-4 py-3 bg-canvas border border-soft rounded-xl text-ink transition-all"
              />
            </div>
            <div>
              <label
                className="block text-[11px] text-muted tracking-[0.18em] uppercase mb-2 font-semibold"
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '0.375rem',
                }}
              >
                <Globe size={12} /> País del cliente{' '}
                <span className="text-magenta">*</span>
              </label>
              <select
                value={clientData.country}
                onChange={(e) =>
                  setClientData({ ...clientData, country: e.target.value })
                }
                className="w-full px-4 py-3 bg-canvas border border-soft rounded-xl text-ink transition-all appearance-none cursor-pointer"
                style={{
                  backgroundImage:
                    "url(\"data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%236B6480' stroke-width='2'%3E%3Cpolyline points='6 9 12 15 18 9'/%3E%3C/svg%3E\")",
                  backgroundRepeat: 'no-repeat',
                  backgroundPosition: 'right 1rem center',
                }}
              >
                {PAISES_LATAM.map((p) => (
                  <option key={p} value={p}>
                    {p}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label className="block text-[11px] text-muted tracking-[0.18em] uppercase mb-2 font-semibold">
                Industria / Sector
              </label>
              <input
                type="text"
                value={clientData.industria}
                onChange={(e) =>
                  setClientData({ ...clientData, industria: e.target.value })
                }
                placeholder="Ej: Retail, Banca, Telecomunicaciones, Salud"
                className="w-full px-4 py-3 bg-canvas border border-soft rounded-xl text-ink transition-all"
              />
              <p className="text-[10px] text-muted mt-1.5 leading-relaxed">
                Opcional. Ayuda a priorizar competencias comunes en tu
                industria.
              </p>
            </div>
            <div>
              <label
                className="block text-[11px] text-muted tracking-[0.18em] uppercase mb-2 font-semibold"
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '0.375rem',
                }}
              >
                <User size={12} /> Tu nombre (CX){' '}
                <span className="text-magenta">*</span>
              </label>
              <input
                type="text"
                value={clientData.cxUser}
                onChange={(e) =>
                  setClientData({ ...clientData, cxUser: e.target.value })
                }
                placeholder="Ej: María González"
                className="w-full px-4 py-3 bg-canvas border border-soft rounded-xl text-ink transition-all"
              />
            </div>
          </div>
        </section>

        {/* Logo */}
        <section className="bg-paper border border-soft rounded-2xl p-8 shadow-brand">
          <div className="flex items-center justify-between gap-2 mb-6">
            <div className="flex items-center gap-2">
              <Palette size={18} className="text-brand" />
              <h3
                className="text-xl text-ink"
                style={{ fontWeight: 700, letterSpacing: '-0.02em' }}
              >
                Logo y branding del cliente
              </h3>
            </div>
            <span className="text-[10px] text-muted tracking-[0.18em] uppercase font-semibold bg-tint px-2.5 py-1 rounded-full">
              Opcional
            </span>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
            <div>
              <input
                ref={fileInputRef}
                type="file"
                accept="image/*"
                onChange={(e) => handleLogoUpload(e.target.files[0])}
                className="hidden"
              />
              <button
                onClick={() => fileInputRef.current?.click()}
                className="w-full h-48 border-2 border-dashed border-soft rounded-xl hover:border-brand transition-all flex flex-col items-center justify-center gap-3 bg-canvas group"
              >
                {logoDataUrl ? (
                  <img
                    src={logoDataUrl}
                    alt="Logo"
                    className="max-h-40 max-w-full object-contain"
                  />
                ) : (
                  <>
                    <Upload
                      size={28}
                      className="text-muted group-hover:text-brand transition-colors"
                      strokeWidth={1.5}
                    />
                    <div className="text-sm text-muted">
                      Haz clic para subir logo
                    </div>
                    <div className="text-xs text-muted">PNG, JPG, SVG</div>
                  </>
                )}
              </button>
            </div>
            <div>
              <div className="text-[11px] text-muted tracking-[0.18em] uppercase mb-3 font-semibold">
                Paleta extraída del logo
              </div>
              {logoColors.length > 0 ? (
                <div className="space-y-2">
                  {logoColors.map((color, i) => (
                    <div key={i} className="flex items-center gap-3">
                      <div
                        className="w-12 h-12 rounded-lg border border-soft shadow-brand"
                        style={{ backgroundColor: color }}
                      />
                      <div>
                        <div className="font-mono text-sm text-ink">
                          {color.toUpperCase()}
                        </div>
                        <div className="text-xs text-muted">
                          {i === 0
                            ? 'Primario'
                            : i === 1
                            ? 'Secundario'
                            : `Acento ${i - 1}`}
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                <div className="text-sm text-muted italic">
                  Sube un logo para extraer los colores del cliente.
                </div>
              )}
            </div>
          </div>
        </section>

        {/* Aviso sobre insumos */}
        <div className="bg-tint border border-soft rounded-xl px-5 py-4 flex items-start gap-3">
          <AlertCircle size={18} className="text-brand flex-shrink-0 mt-0.5" />
          <div className="text-sm text-ink leading-relaxed">
            <span className="font-semibold">
              Los 4 campos siguientes son opcionales
            </span>
            , pero{' '}
            <span className="text-magenta font-semibold">
              al menos uno debe tener contenido
            </span>{' '}
            (texto o archivo adjunto) para poder generar la homologación. Puedes
            combinar texto libre con archivos en un mismo campo.
          </div>
        </div>

        {/* 4 Campos de insumos */}
        {CAMPOS_INSUMOS.map((campo) => (
          <InsumoField
            key={campo.key}
            fieldConfig={campo}
            value={clientData[campo.key]}
            files={clientFiles[campo.key]}
            onTextChange={(v) =>
              setClientData({ ...clientData, [campo.key]: v })
            }
            onFilesAdd={(arr) => addFilesToField(campo.key, arr)}
            onFileRemove={(i) => removeFileFromField(campo.key, i)}
          />
        ))}

        {/* CTA */}
        <div className="flex items-center justify-between pt-4 flex-wrap gap-4">
          <div className="text-xs text-muted">
            <span className="text-magenta">*</span> Campos obligatorios · Al
            menos un insumo (texto o archivo) requerido
          </div>
          <button
            onClick={generateHomologacion}
            disabled={!canSubmit}
            className="gradient-brand text-white px-8 py-4 rounded-full hover-shift disabled:opacity-30 disabled:cursor-not-allowed flex items-center gap-3 shadow-brand-lg font-semibold"
          >
            <Sparkles size={18} />
            <span>Generar homologación</span>
            <ChevronRight size={18} />
          </button>
        </div>
      </div>
    </div>
  );
}

// ============================================================
// PROCESSING VIEW
// ============================================================
function ProcessingView({ progressMsg }) {
  return (
    <div className="max-w-2xl mx-auto py-24 text-center">
      <div className="relative inline-block mb-10">
        <div
          className="absolute inset-0 rounded-full blur-xl opacity-40"
          style={{ background: 'linear-gradient(135deg, #D91B7E, #F5A623)' }}
        />
        <div className="relative w-20 h-20 rounded-full gradient-brand flex items-center justify-center shadow-brand-lg">
          <Loader2
            size={32}
            className="text-white animate-spin"
            strokeWidth={2}
          />
        </div>
      </div>
      <div className="inline-flex items-center gap-2 px-3 py-1 bg-tint rounded-full mb-4">
        <div className="w-1.5 h-1.5 rounded-full bg-magenta animate-pulse" />
        <span className="text-[11px] text-brand tracking-[0.18em] uppercase font-semibold">
          Procesando
        </span>
      </div>
      <h2
        className="font-display text-ink mb-6"
        style={{
          fontSize: 'clamp(1.75rem, 3.5vw, 2.75rem)',
          lineHeight: '1.1',
        }}
      >
        {progressMsg || 'Generando homologación...'}
      </h2>
      <p className="text-muted max-w-md mx-auto leading-relaxed">
        Claude está analizando los insumos del cliente y homologándolos contra
        las 149 competencias del diccionario Evaluar. Esto puede tomar 15-45
        segundos.
      </p>
    </div>
  );
}

// ============================================================
// PREVIEW VIEW (tabla editable)
// ============================================================
function PreviewView({ result, setResult, exportToExcel, resetFlow }) {
  const [editMode, setEditMode] = useState(false);

  const updateHomologacion = (hIndex, eIndex, field, value) => {
    const newResult = { ...result };
    if (eIndex === null) {
      newResult.homologaciones[hIndex][field] = value;
    } else {
      newResult.homologaciones[hIndex].homologacionesEvaluar[eIndex][field] =
        value;
    }
    setResult(newResult);
  };

  const removeEvaluar = (hIndex, eIndex) => {
    const newResult = { ...result };
    newResult.homologaciones[hIndex].homologacionesEvaluar.splice(eIndex, 1);
    setResult(newResult);
  };

  const primary = result.logoColors?.[0] || '#2D3E2F';

  return (
    <div>
      {/* Header del resultado */}
      <div className="mb-10 pb-10 border-b border-soft">
        <div className="flex items-start justify-between gap-8">
          <div className="flex-1">
            <div className="text-xs text-muted tracking-widest uppercase mb-3">
              Propuesta generada
            </div>
            <div className="flex items-end gap-6">
              {result.logoDataUrl && (
                <img
                  src={result.logoDataUrl}
                  alt="Logo"
                  className="h-16 object-contain"
                />
              )}
              <h2
                className="font-display text-ink"
                style={{
                  fontSize: 'clamp(2rem, 4vw, 3.5rem)',
                  lineHeight: '1',
                  letterSpacing: '-0.035em',
                }}
              >
                {result.cliente}
              </h2>
            </div>
            <p className="mt-6 text-lg text-muted max-w-3xl leading-relaxed">
              {result.resumenEjecutivo}
            </p>
            <div className="flex gap-6 mt-6 text-sm text-muted">
              <span>
                Homologado por <span className="text-ink">{result.cxUser}</span>
              </span>
              <span>·</span>
              <span>
                {new Date(result.fecha).toLocaleDateString('es-EC', {
                  year: 'numeric',
                  month: 'long',
                  day: 'numeric',
                })}
              </span>
              <span>·</span>
              <span>
                {result.ejesOrigen === 'cliente'
                  ? 'Ejes del cliente'
                  : 'Ejes propuestos por IA'}
              </span>
            </div>
          </div>
          <div className="flex flex-col gap-2">
            <button
              onClick={() => setEditMode(!editMode)}
              className="px-5 py-2.5 border border-brand rounded-full text-sm flex items-center gap-2 hover-shift text-brand hover:bg-tint font-semibold"
            >
              {editMode ? (
                <>
                  <Check size={14} /> Terminar edición
                </>
              ) : (
                <>
                  <Edit3 size={14} /> Editar
                </>
              )}
            </button>
            <button
              onClick={exportToExcel}
              className="px-5 py-2.5 gradient-brand text-white rounded-full text-sm flex items-center gap-2 hover-shift shadow-brand font-semibold"
            >
              <Download size={14} /> Descargar Excel
            </button>
            <button
              onClick={exportToPDF}
              className="px-5 py-2.5 bg-paper border border-brand text-brand rounded-full text-sm flex items-center gap-2 hover-shift font-semibold hover:bg-tint"
            >
              <Download size={14} /> Descargar PDF
            </button>
            <button
              onClick={resetFlow}
              className="px-5 py-2.5 text-sm flex items-center gap-2 hover-shift text-muted hover:text-brand"
            >
              <Plus size={14} /> Nueva homologación
            </button>
          </div>
        </div>
      </div>

      {/* Ejes */}
      <div className="mb-10">
        <div className="text-[11px] text-muted tracking-[0.18em] uppercase mb-4 font-semibold">
          Ejes de competencias
        </div>
        <div className="flex flex-wrap gap-2">
          {(result.ejesPropuestos || []).map((eje, i) => (
            <span
              key={i}
              className="px-4 py-2 text-sm rounded-full font-medium shadow-brand"
              style={{ backgroundColor: primary, color: 'white' }}
            >
              {eje}
            </span>
          ))}
        </div>
      </div>

      {/* Tabla de homologación */}
      <div className="mb-10">
        <div className="text-[11px] text-muted tracking-[0.18em] uppercase mb-4 font-semibold">
          Matriz de homologación
        </div>
        <div className="bg-paper border border-soft rounded-2xl overflow-x-auto shadow-brand">
          <table className="w-full text-sm">
            <thead>
              <tr
                className="border-b border-soft"
                style={{ backgroundColor: '#F3EFFB' }}
              >
                <th className="px-3 py-3 text-left font-medium text-ink text-xs">
                  Eje cliente
                </th>
                <th className="px-3 py-3 text-left font-medium text-ink text-xs">
                  Competencia cliente
                </th>
                <th className="px-3 py-3 text-left font-medium text-ink text-xs">
                  Definición cliente
                </th>
                <th className="px-3 py-3 text-center font-medium text-ink text-xs">
                  Nivel de cargo
                </th>
                <th className="px-3 py-3 text-left font-medium text-ink text-xs">
                  Eje Evaluar
                </th>
                <th className="px-3 py-3 text-left font-medium text-ink text-xs">
                  Competencia Evaluar
                </th>
                <th className="px-3 py-3 text-left font-medium text-ink text-xs">
                  Definición Evaluar
                </th>
                <th className="px-3 py-3 text-center font-medium text-ink text-xs">
                  Nivel de desarrollo
                </th>
                <th className="px-3 py-3 text-left font-medium text-ink text-xs">
                  Justificación
                </th>
                {editMode && <th className="w-10"></th>}
              </tr>
            </thead>
            <tbody>
              {(result.homologaciones || []).map((h, hIdx) => {
                // Soporte para nueva estructura (homologacionesPorNivel) y legacy (homologacionesEvaluar)
                const evaluarList =
                  h.homologacionesPorNivel || h.homologacionesEvaluar || [];
                return evaluarList.map((e, eIdx) => {
                  const nivel =
                    e.nivelDesarrolloExigido ?? e.nivelDesarrollo ?? 5;
                  const ejeEv = e.ejeEvaluar || e.categoriaEvaluar || '';
                  const nivelCargo = e.nivelCargo || h.nivelCargoCliente || '';
                  return (
                    <tr
                      key={`${hIdx}-${eIdx}`}
                      className="border-b border-soft last:border-b-0 align-top text-xs"
                    >
                      {/* Eje cliente */}
                      {eIdx === 0 && (
                        <td
                          rowSpan={evaluarList.length}
                          className="px-3 py-4 border-r border-soft"
                          style={{ backgroundColor: '#FAF7F2' }}
                        >
                          {editMode ? (
                            <input
                              value={h.ejeCliente}
                              onChange={(ev) =>
                                updateHomologacion(
                                  hIdx,
                                  null,
                                  'ejeCliente',
                                  ev.target.value
                                )
                              }
                              className="text-xs text-ink uppercase tracking-widest w-full bg-transparent border-b border-soft pb-1"
                            />
                          ) : (
                            <div className="text-xs text-ink uppercase tracking-widest">
                              {h.ejeCliente}
                            </div>
                          )}
                        </td>
                      )}
                      {/* Competencia cliente */}
                      {eIdx === 0 && (
                        <td
                          rowSpan={evaluarList.length}
                          className="px-3 py-4 border-r border-soft"
                          style={{ backgroundColor: '#FAF7F2' }}
                        >
                          {editMode ? (
                            <div className="space-y-1">
                              <input
                                value={h.competenciaCliente}
                                onChange={(ev) =>
                                  updateHomologacion(
                                    hIdx,
                                    null,
                                    'competenciaCliente',
                                    ev.target.value
                                  )
                                }
                                className="font-medium text-ink w-full bg-transparent border-b border-soft pb-1"
                              />
                              <input
                                value={h.nivelCargoCliente || ''}
                                onChange={(ev) =>
                                  updateHomologacion(
                                    hIdx,
                                    null,
                                    'nivelCargoCliente',
                                    ev.target.value
                                  )
                                }
                                className="text-xs text-muted bg-transparent border-b border-soft w-full"
                                placeholder="Nivel del cargo (legacy)"
                              />
                            </div>
                          ) : (
                            <>
                              <div className="font-medium text-ink">
                                {h.competenciaCliente}
                              </div>
                              {h.fraseCliente && (
                                <div className="text-[10px] italic text-muted mt-1">
                                  {h.fraseCliente}
                                </div>
                              )}
                              {h.nivelCargoCliente && (
                                <div className="text-[10px] text-muted mt-2 uppercase tracking-wider">
                                  {h.nivelCargoCliente}
                                </div>
                              )}
                            </>
                          )}
                        </td>
                      )}
                      {/* Definición cliente */}
                      {eIdx === 0 && (
                        <td
                          rowSpan={evaluarList.length}
                          className="px-3 py-4 border-r border-soft text-muted leading-relaxed"
                          style={{ backgroundColor: '#FAF7F2' }}
                        >
                          {editMode ? (
                            <textarea
                              value={h.definicionCliente || ''}
                              onChange={(ev) =>
                                updateHomologacion(
                                  hIdx,
                                  null,
                                  'definicionCliente',
                                  ev.target.value
                                )
                              }
                              className="w-full bg-transparent text-xs min-h-24"
                              placeholder="— Sin definición del cliente —"
                            />
                          ) : (
                            h.definicionCliente || (
                              <span className="italic text-muted">
                                — Sin definición —
                              </span>
                            )
                          )}
                        </td>
                      )}
                      {/* NUEVA COLUMNA: Nivel de cargo */}
                      <td className="px-3 py-4 text-center bg-tint">
                        {editMode ? (
                          <input
                            value={nivelCargo}
                            onChange={(ev) =>
                              updateHomologacion(
                                hIdx,
                                eIdx,
                                'nivelCargo',
                                ev.target.value
                              )
                            }
                            className="text-xs text-ink uppercase tracking-wider w-full bg-transparent text-center border-b border-soft"
                            placeholder="Ej: Ejecutor"
                          />
                        ) : (
                          <div className="text-[10px] text-ink uppercase tracking-wider font-semibold">
                            {nivelCargo || '—'}
                          </div>
                        )}
                      </td>
                      {/* Eje Evaluar */}
                      <td className="px-3 py-4">
                        {editMode ? (
                          <input
                            value={ejeEv}
                            onChange={(ev) =>
                              updateHomologacion(
                                hIdx,
                                eIdx,
                                'ejeEvaluar',
                                ev.target.value
                              )
                            }
                            className="text-xs text-muted uppercase tracking-widest w-full bg-transparent border-b border-soft"
                          />
                        ) : (
                          <div className="text-[10px] text-muted uppercase tracking-widest font-semibold">
                            {ejeEv}
                          </div>
                        )}
                      </td>
                      {/* Competencia Evaluar */}
                      <td className="px-3 py-4">
                        {editMode ? (
                          <input
                            value={e.competenciaEvaluar}
                            onChange={(ev) =>
                              updateHomologacion(
                                hIdx,
                                eIdx,
                                'competenciaEvaluar',
                                ev.target.value
                              )
                            }
                            className="font-medium text-ink w-full bg-transparent border-b border-soft"
                          />
                        ) : (
                          <div className="font-medium text-ink">
                            {e.competenciaEvaluar}
                          </div>
                        )}
                      </td>
                      {/* Definición Evaluar */}
                      <td className="px-3 py-4 text-muted leading-relaxed italic">
                        {editMode ? (
                          <textarea
                            value={e.definicionEvaluar}
                            onChange={(ev) =>
                              updateHomologacion(
                                hIdx,
                                eIdx,
                                'definicionEvaluar',
                                ev.target.value
                              )
                            }
                            className="w-full bg-transparent text-xs min-h-20"
                          />
                        ) : (
                          e.definicionEvaluar
                        )}
                      </td>
                      {/* Nivel desarrollo */}
                      <td className="px-3 py-4 text-center">
                        {editMode ? (
                          <input
                            type="number"
                            min="5"
                            max="9"
                            value={nivel}
                            onChange={(ev) =>
                              updateHomologacion(
                                hIdx,
                                eIdx,
                                'nivelDesarrolloExigido',
                                parseInt(ev.target.value) || 5
                              )
                            }
                            className="w-12 text-center font-display text-3xl bg-transparent"
                          />
                        ) : (
                          <span
                            className="inline-flex items-center justify-center w-11 h-11 rounded-xl text-white shadow-brand"
                            style={{
                              background:
                                nivel >= 8
                                  ? 'linear-gradient(135deg, #2E1C5A 0%, #5B2B9C 100%)'
                                  : nivel >= 7
                                  ? 'linear-gradient(135deg, #D91B7E 0%, #E94BA1 100%)'
                                  : 'linear-gradient(135deg, #F5A623 0%, #F8BE5B 100%)',
                              fontWeight: 800,
                              fontSize: '1.35rem',
                              letterSpacing: '-0.02em',
                            }}
                          >
                            {nivel}
                          </span>
                        )}
                      </td>
                      {/* Justificación */}
                      <td className="px-3 py-4 text-muted leading-relaxed">
                        {editMode ? (
                          <textarea
                            value={e.justificacion}
                            onChange={(ev) =>
                              updateHomologacion(
                                hIdx,
                                eIdx,
                                'justificacion',
                                ev.target.value
                              )
                            }
                            className="w-full bg-transparent text-xs min-h-20"
                          />
                        ) : (
                          e.justificacion
                        )}
                      </td>
                      {editMode && (
                        <td className="px-2">
                          <button
                            onClick={() => removeEvaluar(hIdx, eIdx)}
                            className="text-red-500 hover:text-red-700"
                          >
                            <Trash2 size={14} />
                          </button>
                        </td>
                      )}
                    </tr>
                  );
                });
              })}
            </tbody>
          </table>
        </div>
      </div>

      {/* Niveles por familia */}
      {result.nivelesPorFamilia && result.nivelesPorFamilia.length > 0 && (
        <div className="mb-10">
          <div className="text-[11px] text-muted tracking-[0.18em] uppercase mb-4 font-semibold">
            Niveles por familia de cargo
          </div>
          <div className="grid grid-cols-2 gap-4">
            {result.nivelesPorFamilia.map((n, i) => (
              <div
                key={i}
                className="bg-paper border border-soft rounded-2xl p-5 shadow-brand hover-shift"
              >
                <div className="flex items-center justify-between mb-3">
                  <div
                    className="text-ink"
                    style={{
                      fontWeight: 700,
                      fontSize: '1.15rem',
                      letterSpacing: '-0.02em',
                    }}
                  >
                    {n.familia}
                  </div>
                  <div className="flex items-center gap-1">
                    <span
                      className="px-3 py-1 text-xs rounded-full font-semibold text-white"
                      style={{
                        background:
                          'linear-gradient(135deg, #2E1C5A 0%, #D91B7E 100%)',
                      }}
                    >
                      {n.rangoMinimo}–{n.rangoMaximo}
                    </span>
                  </div>
                </div>
                {n.cargosDetectados && n.cargosDetectados.length > 0 && (
                  <div className="text-xs text-muted leading-relaxed">
                    {n.cargosDetectados.join(' · ')}
                  </div>
                )}
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Comparativa de mercado */}
      {result.comparativaMercado && (
        <div className="mb-10">
          <div className="flex items-center gap-2 mb-4">
            <TrendingUp size={16} className="text-brand" />
            <div className="text-[11px] text-muted tracking-[0.18em] uppercase font-semibold">
              Comparativa de mercado — {result.comparativaMercado.industria}
            </div>
          </div>
          <div className="bg-paper border border-soft rounded-2xl p-6 shadow-brand">
            {/* Competencias comunes identificadas */}
            {result.comparativaMercado.competenciasComunesIdentificadas &&
              result.comparativaMercado.competenciasComunesIdentificadas
                .length > 0 && (
                <div className="mb-5">
                  <div className="text-xs text-muted uppercase tracking-wider mb-2 font-semibold">
                    Competencias más comunes en la industria
                  </div>
                  <div className="flex flex-wrap gap-2">
                    {result.comparativaMercado.competenciasComunesIdentificadas.map(
                      (comp, i) => (
                        <span
                          key={i}
                          className="px-3 py-1.5 bg-tint text-ink text-xs rounded-full border border-soft"
                        >
                          {comp}
                        </span>
                      )
                    )}
                  </div>
                </div>
              )}
            {/* Análisis */}
            {result.comparativaMercado.analisis && (
              <div className="mb-5">
                <div className="text-xs text-muted uppercase tracking-wider mb-2 font-semibold">
                  Análisis comparativo
                </div>
                <div className="text-sm text-ink leading-relaxed whitespace-pre-line">
                  {result.comparativaMercado.analisis}
                </div>
              </div>
            )}
            {/* Recomendaciones */}
            {result.comparativaMercado.recomendaciones &&
              result.comparativaMercado.recomendaciones.length > 0 && (
                <div>
                  <div className="text-xs text-muted uppercase tracking-wider mb-2 font-semibold">
                    Recomendaciones
                  </div>
                  <ul className="space-y-2">
                    {result.comparativaMercado.recomendaciones.map((rec, i) => (
                      <li
                        key={i}
                        className="flex gap-2 text-sm text-ink leading-relaxed"
                      >
                        <span className="text-brand mt-1">•</span>
                        <span>{rec}</span>
                      </li>
                    ))}
                  </ul>
                </div>
              )}
          </div>
        </div>
      )}
    </div>
  );
}

// ============================================================
// HISTORY VIEW
// ============================================================
function HistoryView({ history, setResult, setStep }) {
  if (history.length === 0) {
    return (
      <div className="max-w-2xl mx-auto py-24 text-center">
        <div className="w-20 h-20 rounded-full bg-tint mx-auto mb-8 flex items-center justify-center">
          <Archive size={32} className="text-brand" strokeWidth={1.5} />
        </div>
        <h2
          className="font-display text-ink mb-4"
          style={{
            fontSize: 'clamp(1.75rem, 3.5vw, 2.5rem)',
            letterSpacing: '-0.035em',
          }}
        >
          Repositorio vacío
        </h2>
        <p className="text-muted mb-8">
          Aquí aparecerán las homologaciones que guardes al descargar.
        </p>
        <button
          onClick={() => setStep('new')}
          className="gradient-brand text-white px-6 py-3 rounded-full hover-shift inline-flex items-center gap-2 shadow-brand-lg font-semibold"
        >
          <Sparkles size={16} /> Crear la primera
        </button>
      </div>
    );
  }

  return (
    <div>
      <div className="mb-10">
        <div className="inline-flex items-center gap-2 px-3 py-1 bg-tint rounded-full mb-4">
          <span className="text-[11px] text-brand tracking-[0.18em] uppercase font-semibold">
            Repositorio interno
          </span>
        </div>
        <div className="flex items-end justify-between">
          <h2
            className="font-display text-ink"
            style={{
              fontSize: 'clamp(2rem, 4vw, 3.5rem)',
              letterSpacing: '-0.035em',
            }}
          >
            {history.length}{' '}
            <span className="gradient-text">homologaciones</span>
          </h2>
          <div className="text-sm text-muted">
            En la versión web final, podrás descargar un consolidado en Excel.
          </div>
        </div>
      </div>

      <div className="bg-paper border border-soft rounded-2xl overflow-hidden shadow-brand">
        <table className="w-full text-sm">
          <thead
            className="border-b border-soft"
            style={{ backgroundColor: '#F3EFFB' }}
          >
            <tr>
              <th className="px-5 py-3 text-left font-medium text-ink">
                Cliente
              </th>
              <th className="px-5 py-3 text-left font-medium text-ink">
                Fecha
              </th>
              <th className="px-5 py-3 text-left font-medium text-ink">
                Responsable CX
              </th>
              <th className="px-5 py-3 text-center font-medium text-ink">
                Competencias
              </th>
              <th className="px-5 py-3 text-left font-medium text-ink">Ejes</th>
              <th className="px-5 py-3"></th>
            </tr>
          </thead>
          <tbody>
            {history.map((h) => (
              <tr
                key={h.id}
                className="border-b border-soft last:border-b-0 hover:bg-tint transition-colors"
              >
                <td className="px-5 py-4">
                  <div className="flex items-center gap-3">
                    {h.logoDataUrl && (
                      <img
                        src={h.logoDataUrl}
                        alt=""
                        className="w-8 h-8 object-contain"
                      />
                    )}
                    <div className="font-medium text-ink">{h.cliente}</div>
                  </div>
                </td>
                <td className="px-5 py-4 text-muted">
                  {new Date(h.fecha).toLocaleDateString('es-EC', {
                    day: '2-digit',
                    month: 'short',
                    year: 'numeric',
                  })}
                </td>
                <td className="px-5 py-4 text-muted">{h.cxUser}</td>
                <td className="px-5 py-4 text-center text-muted">
                  {h.homologaciones?.length || 0}
                </td>
                <td className="px-5 py-4 text-muted text-xs">
                  {(h.ejesPropuestos || []).slice(0, 2).join(' · ')}
                  {h.ejesPropuestos?.length > 2 && '…'}
                </td>
                <td className="px-5 py-4 text-right">
                  <button
                    onClick={() => {
                      setResult(h);
                      setStep('preview');
                    }}
                    className="text-sm text-brand hover:text-magenta flex items-center gap-1 ml-auto font-semibold transition-colors"
                  >
                    <Eye size={14} /> Ver
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
