/* ==========================================================
   IMEN · M.Sc. Semiconductor Engineering
   Shared configuration + content + bilingual UI strings
   ========================================================== */

const APP_CONFIG = {
  // Paste the deployed Google Apps Script Web App URL here when ready.
  SCRIPT_URL: "https://script.google.com/macros/s/AKfycbyVrdk1_LxoJn4RfU7ImQZ_CMJxDdp_nYld-T7pxj7ZduaPrimt4yCSYVJhWE3H6LG6/exec",
  // Runs on localStorage until SCRIPT_URL is filled in.
  DEMO_MODE: false,
  // Demo-only dashboard password (live mode verifies server-side).
  DEMO_PASSWORD: "imen-demo",
  RESPONSE_KEY: "msc_semicond_responses_v1",
  LANG_KEY: "msc_semicond_lang",
  READ_KEY: "msc_semicond_read_v1",
  DRAFT_KEY: "msc_semicond_draft_v1"
};

/* ------------------------------------------------------------------
   SURVEY CONTENT — taken verbatim from the official questionnaire.
   Only obvious typos / spacing have been corrected.
------------------------------------------------------------------ */
const PROGRAMME = {
  name: "Master of Science in Semiconductor Engineering",
  short: "M.Sc. Semiconductor Engineering",
  institute: "Institute of Microengineering and Nanoelectronics (IMEN)",
  university: "Universiti Kebangsaan Malaysia (UKM)",
  intro:
    "This questionnaire aims to evaluate and ensure the marketability of the new academic programme proposed by the Institute of Microengineering and Nanoelectronics (IMEN), UKM, namely the Master of Science in Semiconductor Engineering by coursework. The Institute of Microengineering and Nanoelectronics at Universiti Kebangsaan Malaysia intends to offer a study programme to build a human capital workforce capable in technologies related to microengineering and nanotechnology through research and development activities in MEMS technology, VLSI design, sensor technology, advanced materials, packaging technology, photonics, up to marketable prototype-level developments. In addition, the findings of this study are important to ensure that the programme structure complies with the standards of the Malaysian Qualifications Agency (MQA) and supports national aspirations such as the National Semiconductor Strategy (NSS) and the New Industrial Master Plan 2030 (NIMP 2030) in producing a highly skilled workforce. Therefore, IMEN cordially invites YBhg. Prof. / Dato' / Datuk / Dr. / Mr. / Mrs. / Ms. to spare some time to participate in this survey, as the findings can help improve the development of this programme.",
  mode: "Full-time (12 months) / Part-time (24 months)",
  credits: "40 credits",
  split: "Core 34 credits · Elective 6 credits"
};

const CATEGORY_QUESTION =
  "Respondent Category / Employment Sector:";

const CATEGORIES = [
  "Semiconductor / Electronics Industry (MNC / SME)",
  "Design Company / Design House (IC Design)",
  "Assembly, Test & Packaging (OSAT) / Wafer Fab",
  "Academia / Lecturer / Researcher",
  "Civil Servant / Government Agency / Sponsorship Agency (e.g., PTPTN, JPA, MIDA)",
  "Government-Linked Company (GLC)",
  "Private Company – Employer / Executive / Staff",
  "Self-Employed / Freelancer / Business Owner",
  "Student",
  "Others"
];

const CRITERIA_QUESTION =
  "What are the key criteria that make this programme valuable to the market? (Multiple selections allowed)";

const CRITERIA = [
  "Balance between fundamental theory and practical/cleanroom training",
  "Exposure to cutting-edge technologies (Advanced Packaging, MEMS, IC Design)",
  "Availability of industry-based projects / industrial problem-solving",
  "Fulfillment of critical skill gap requirements (Talent Gap) in the local industry"
];

const CORE_QUESTION =
  "Please rate the suitability and importance of the following proposed courses for the Master of Science in Semiconductor Engineering programme: 1. Core Courses (34 Credits)";

const ELECTIVE_QUESTION =
  "2. Elective Courses (Choose 6 Credits)";

const SCALE_LABELS = {
  1: "Highly Unsuitable",
  2: "Unsuitable",
  3: "Moderate",
  4: "Suitable",
  5: "Highly Suitable"
};

const SCALE_TEXT = "Scale: 1 = Highly Unsuitable | 2 = Unsuitable | 3 = Moderate | 4 = Suitable | 5 = Highly Suitable";

const COURSES = [
  {
    id: "core_1",
    type: "core",
    title: "Semiconductor Device Physics",
    summary:
      "Examines the principles of quantum physics and the properties of semiconductor materials, as well as the operation of microelectronic devices.",
    overview:
      "Studies the fundamental principles of quantum physics and solid-state physics to understand the electrical, optical, and structural properties of semiconductor materials. This course covers the analysis of charge carrier transport mechanisms, energy band formation, and p-n junctions that form the foundation of microelectronic device operations. Students will learn the operating theory of key devices such as diodes, bipolar junction transistors (BJTs), and metal-oxide-semiconductor field-effect transistors (MOSFETs). Emphasis is placed on understanding physical phenomena at the micro and nano scales that affect device performance. Through this understanding, students can analyze device C-V/I-V characteristics and design next-generation microelectronic devices that are more efficient. This course provides a solid theoretical foundation to support fabrication processing, integrated circuit design, and chip packaging in the semiconductor industry.",
    scene: "device",
    icon: "M12 3v18M5 8h14M5 16h14M8 5l8 14"
  },
  {
    id: "core_2",
    type: "core",
    title: "Microfabrication Technology & Cleanroom",
    summary:
      "Studies microfabrication processes, wafer processing, and the operational environment of cleanrooms.",
    overview:
      "Provides comprehensive exposure to the principles and techniques of wafer processing in microelectronic device manufacturing. This course covers key steps in microfabrication such as lithography, etching, thin-film deposition (CVD/PVD), diffusion, and ion implantation. Additionally, students will learn cleanroom operating protocols, contamination control, and the semiconductor chip manufacturing process flow from early stages to completion. Emphasis is also placed on operating high-tech processing equipment and parameter optimization to improve production yield. Through a combination of theory and practical training in cleanroom facilities, students will develop technical skills relevant to industry standards. This course is vital in preparing a highly skilled workforce for the global semiconductor fabrication sector.",
    scene: "fab",
    icon: "M4 16h16M7 16V8l5-3 5 3v8M9 16v-4h6v4"
  },
  {
    id: "core_3",
    type: "core",
    title: "Semiconductor Materials & Characterization",
    summary:
      "Focuses on the analysis of physical properties, structure, and characterization techniques of semiconductor materials.",
    overview:
      "Focuses on an in-depth analysis of the physical, electrical, optical, and structural properties of basic and advanced semiconductor materials such as Silicon, GaAs, and wide bandgap materials (GaN/SiC). Students will learn physical and structural characterization techniques to identify crystal quality, impurities, and material defects. The course covers key analytical methods such as X-Ray Diffraction (XRD), Scanning Electron Microscopy (SEM), Atomic Force Microscopy (AFM), and C-V/I-V electrical testing. Emphasis is placed on the relationship between material microstructure and semiconductor device performance. Through this understanding, students can evaluate material suitability for specific applications and optimize fabrication processes. This course provides fundamental expertise essential for research and development (R&D) as well as material quality control in the local and international semiconductor engineering industry.",
    scene: "materials",
    icon: "M12 3l7 4v10l-7 4-7-4V7l7-4zM12 3v18M5 7l14 10"
  },
  {
    id: "core_4",
    type: "core",
    title: "Assembly, Test & Packaging",
    summary:
      "Covers chip packaging techniques, device testing methods, and back-end processing integration.",
    overview:
      "Provides a deep understanding of back-end processes in semiconductor device manufacturing. This course covers chip packaging techniques from traditional methods to advanced technologies such as 2.5D/3D packaging, Flip-Chip, Wafer-Level Packaging, and heterogeneous integration. Students will learn device testing methods, including Wafer-Level Testing and Automated Test Equipment (ATE), to ensure chip functionality and quality. Additionally, this course delves into thermal management issues, signal integrity, and mechanical packaging reliability. Emphasis is placed on assembly process optimization to support advanced packaging trends in the semiconductor industry. This course produces graduates competent in addressing system integration challenges and global chip product quality assurance.",
    scene: "package",
    icon: "M4 8l8-4 8 4-8 4-8-4zM4 8v8l8 4 8-4V8M12 12v8"
  },
  {
    id: "core_5",
    type: "core",
    title: "Microelectronic Systems and Applications",
    summary:
      "Explores the design of microelectronic systems and their applications in information technology and industry.",
    overview:
      "Explores the principles of design and integration of advanced microelectronic systems and their role in driving information technology and industrial applications. This course covers microelectronic architecture, signal processors, System-on-Chip (SoC) architecture, and micro power devices. Students will learn how to align microelectronic components with practical applications such as artificial intelligence (AI), data centers, 5G/6G communications, medical devices, and electric automotive systems. Emphasis is placed on system performance optimization, energy efficiency, and device miniaturization to meet current industry demands. Through this course, students will be capable of designing and analyzing complex, highly innovative microelectronic systems, producing experts ready to support the growth of the local and international high-tech ecosystem.",
    scene: "systems",
    icon: "M6 6h12v12H6zM9 9h6v6H9zM12 3v3M12 18v3M3 12h3M18 12h3"
  },
  {
    id: "core_6",
    type: "core",
    title: "Research Methodology",
    summary:
      "Trains skills in systematically planning, executing, and analyzing scientific studies and engineering experiments.",
    overview:
      "Equips students with fundamental skills and systematic methodologies to plan, execute, and analyze scientific research in the field of semiconductor engineering. This course covers research problem formulation techniques, critical literature review, Design of Experiments (DoE), and quantitative data collection and analysis using advanced software tools. Students will learn research ethics management, copyright and intellectual property compliance, and academic writing techniques for high-impact journal publications. Emphasis is given to industry-oriented problem-solving methods and the preparation of structured research proposals for final projects. Through this course, students are trained to think scientifically and analytically in managing complex semiconductor R&D projects. This course is essential to ensure research outcomes meet local and international academic and industrial qualification standards.",
    scene: "research",
    icon: "M10 3h4M11 3v6l-5 9a2 2 0 002 3h8a2 2 0 002-3l-5-9V3M8 15h8"
  },
  {
    id: "core_7",
    type: "core",
    title: "Project Management",
    summary:
      "Teaches principles of scheduling, resource management, budgeting, and risk control in semiconductor projects.",
    overview:
      "Equips students with leadership skills and strategic governance specifically tailored for the high-tech semiconductor industry. This course teaches principles of project scheduling, resource planning and allocation, budgeting, and detailed risk management. Students will learn standard project management frameworks such as Agile and Waterfall, based on PMBOK methodology adapted to the complex chip development lifecycle. Emphasis is placed on quality control, R&D timeline compliance, stakeholder communication, and high-cost fabrication facility management. Through real semiconductor industry case studies, students are trained to resolve operational issues and efficiently lead post-project activities. This course produces lead engineers capable of managing complex microelectronics projects according to international industry standards.",
    scene: "management",
    icon: "M4 19h16M7 19v-6M12 19V8M17 19v-9M5 6l4 2 4-3 5 2"
  },
  {
    id: "core_8",
    type: "core",
    title: "Project I (Final Project)",
    summary:
      "The initiation phase of the final research project: proposal development, literature review, and the start of experimental work.",
    overview:
      "Project I — Represents the initiation phase of a major research component oriented toward application and semiconductor industry problem-solving. This course covers comprehensive project proposal development, critical literature review, and determination of appropriate research methodologies. Students will begin the initial stages of experimental execution, simulation design, or characterization work in cleanrooms and specialized laboratories. This course emphasizes critical thinking, high-impact problem-solving, and aligning projects with local and international semiconductor ecosystem needs. Under the guidance of academic supervisors and industry partners, students refine analytical skills and present preliminary research progress. Project I builds a strong practical foundation to be completed in the subsequent phase.",
    scene: "project1",
    icon: "M4 18h5V6H4zM10 18h5V9h-5zM16 18h5v-7h-5zM7 4l5 3 6-1"
  },
  {
    id: "core_9",
    type: "core",
    title: "Project II (Final Project)",
    summary:
      "The continuation and culmination of the research initiated in Project I: advanced experiments, analysis, dissertation, and oral defence.",
    overview:
      "Project II — Serves as the continuation and culmination of research initiated in Project I, focusing on solving key industrial semiconductor research problems. This course covers advanced experimental execution, laboratory simulation, device fabrication, and in-depth data collection. Students will perform high-performance data analysis, modelling, and synthesis of research findings to produce innovative solutions to real-world technical problems. Emphasis is placed on process optimization, validation of experimental results, and technology feasibility assessment. Research findings must be presented as a high-standard scientific dissertation and delivered via an oral presentation to an evaluation panel. This course empowers graduates' higher-order thinking skills as caliber researchers capable of contributing directly to the development of the nation's microengineering and nanotechnology ecosystem.",
    scene: "project2",
    icon: "M4 18h5V9H4zM10 18h5V6h-5zM16 18h5v-9h-5zM7 4l5 3 6-1"
  },
  {
    id: "elective_1",
    type: "elective",
    title: "Integrated Circuit (VLSI) Design and IoT",
    summary:
      "Studies micro-based circuit design and its integration into Internet-of-Things applications.",
    overview:
      "Explores the design principles of Very Large Scale Integrated (VLSI) circuits and their integration into the Internet-of-Things (IoT) ecosystem. This course covers micro-based circuit architecture, mixed analogue/digital circuit design, and low-power circuit modelling methodologies. Students will learn to use Computer-Aided Design (CAD) tools and Electronic Design Automation (EDA) modelling to design energy-efficient and compact chips. Emphasis is placed on advanced chip design challenges such as power management, signal integrity, and edge computing for IoT sensor nodes. Through this course, students will be capable of designing VLSI chips optimized for smart applications in automotive, medical, and communications industries, providing critical expertise to support local IC design development.",
    scene: "vlsi",
    icon: "M5 5h14v14H5zM8 8h3v3H8zM13 8h3v3h-3zM8 13h3v3H8zM13 13h3v3h-3z"
  },
  {
    id: "elective_2",
    type: "elective",
    title: "MEMS & Sensor Technology",
    summary:
      "Examines the design, development, and fabrication of micro-electro-mechanical systems and sensors.",
    overview:
      "Examines design, development, and fabrication principles of Micro-Electro-Mechanical Systems (MEMS) and various advanced sensor types. This course covers specialized microfabrication techniques such as bulk micromachining, surface micromachining, and wafer bonding to produce micro-scale devices. Students will learn physical, chemical, and biological sensing mechanisms, including pressure sensors, accelerometers, gyroscopes, and biosensors for medical and industrial applications. Emphasis is placed on structural modeling, sensor integration with electronic circuits, and wafer-level packaging techniques. Through this course, students will be able to design and analyze the performance of high-precision, low-power MEMS sensors, providing practical skills to support the growth of the smart sensor, automotive, and wearable device industries within the semiconductor ecosystem.",
    scene: "mems",
    icon: "M4 15h9l5-6h3M13 15v4M8 7h4v4H8zM3 19h18"
  },
  {
    id: "elective_3",
    type: "elective",
    title: "Photonics & Optoelectronics",
    summary:
      "Explores light signal processing, optoelectronic devices, and photonic integration on-chip.",
    overview:
      "Explores fundamental principles of light signal processing and its interaction with semiconductor materials for developing advanced optoelectronic devices. This course covers the theory and operation of key devices such as light-emitting diodes (LEDs), semiconductor lasers, photodetectors, and solar cells. Students will learn concepts of silicon photonics and direct integration of optical components into photonic integrated circuits (PICs). Emphasis is placed on optimizing light emission efficiency, optical waveguide modes, high-speed signal modulation, and reducing optical power loss. Applications covered include high-speed fiber-optic communication systems, data centers, 5G/6G communication, and industrial optical sensing. Through this course, students are equipped with critical expertise to design and integrate advanced photonic devices to meet future data connectivity needs.",
    scene: "photonics",
    icon: "M3 12c3-6 6-6 9 0s6 6 9 0M3 16h18M12 3v5"
  },
  {
    id: "elective_4",
    type: "elective",
    title: "Failure Analysis & Reliability",
    summary:
      "Studies defect identification techniques, lifetime testing, and device reliability enhancement.",
    overview:
      "Provides detailed understanding of semiconductor device failure mechanisms, defect identification methods, and device reliability enhancement techniques. This course covers reliability testing methods such as Accelerated Life Testing (ALT), High-Temperature Operating Life (HTOL), and thermal shock testing to predict device operational lifespan. Students will learn laboratory-based physical and chemical failure analysis techniques, including Scanning Electron Microscopy (SEM), Emission Microscopy (EMMI), and Focused Ion Beam (FIB) milling to pinpoint root causes of failure during fabrication or packaging stages. Emphasis is placed on process improvement strategies, quality control, and Design for Reliability (DfR). This course produces competent engineers capable of ensuring semiconductor chip quality and durability according to international standards.",
    scene: "failure",
    icon: "M12 3l8 4v5c0 5-3.5 8-8 9-4.5-1-8-4-8-9V7l8-4zM9 10l3 3 3-3M9 15h6"
  },
  {
    id: "elective_5",
    type: "elective",
    title: "Semiconductor Supply Chain Management",
    summary:
      "Studies logistics strategies, production planning, and supply chain dynamics in the semiconductor industry.",
    overview:
      "Examines supply chain dynamics, production planning, and logistics strategies in the complex and global semiconductor industry. This course covers wafer inventory management, fabrication (fab) scheduling, production capacity planning, and supply chain disruption risk mitigation strategies. Students will study geopolitical analysis, raw material supplier management (such as high-purity gases and substrates), and front-end and back-end processing logistics. Emphasis is placed on applying data analytics tools to forecast market demand and optimize chip value chain efficiency. Through this course, students will be able to formulate resilient, cost-effective supply chain strategies, building critical management skills to ensure operational continuity and reliability in the national semiconductor industry.",
    scene: "supply",
    icon: "M12 3a9 9 0 100 18 9 9 0 000-18zM3 12h18M12 3c3 3.5 4.5 6.5 4.5 9S15 17.5 12 21c-3-3.5-4.5-6.5-4.5-9S9 6.5 12 3z"
  }
];

const STUDY_MODE_QUESTION = "Recommended Study Mode Delivery:";
const STUDY_MODES = [
  "Fully Physical Classes",
  "Fully Online",
  "Hybrid (Combination of physical & online)",
  "Executive / Flexible Mode (Weekend classes)"
];

const FEE_QUESTION =
  "Suggested Reasonable Tuition Fee Rate (Local / Citizen):";
const FEE_RANGES = [
  "RM 10,000 - RM 14,999",
  "RM 15,000 - RM 19,999",
  "RM 20,000 - RM 25,000",
  "Others"
];

const INTEREST_QUESTION =
  "Are you interested in enrolling or encouraging staff / colleagues / family members to enroll in this programme?";
const INTEREST_OPTIONS = [
  "Yes, interested in enrolling personally",
  "Yes, interested in nominating/sponsoring company staff",
  "Yes, will encourage family members / friends",
  "Not interested",
  "Unsure"
];

const FEEDBACK_QUESTION =
  "Suggestions for improvement to enhance the development quality of the Master of Science in Semiconductor Engineering programme at IMEN, UKM, such as adding or removing skills/courses? (Open feedback)";

/* ------------------------------------------------------------------
   BILINGUAL UI (interface chrome only — questions stay verbatim)
------------------------------------------------------------------ */
const I18N = {
  en: {
    brand: "IMEN · UKM",
    nav_about: "About",
    nav_curriculum: "Curriculum",
    nav_survey: "Survey",
    nav_lang: "BM",
    hero_eyebrow: "Marketability Study · New Programme",
    hero_title_1: "Master of Science in",
    hero_title_accent: "Semiconductor Engineering",
    hero_lead:
      "Help shape a programme built for the National Semiconductor Strategy. Your perspective will ensure the curriculum is industry-ready, MQA-compliant, and future-focused.",
    hero_cta_start: "Begin the survey",
    hero_cta_explore: "Explore the curriculum",
    hero_meta_mode: "Study mode",
    hero_meta_credits: "Total credits",
    hero_meta_split: "Structure",
    hero_core_elective: "Core 34 · Elective 6",
    hero_scroll: "Scroll to explore",
    stat_1_k: "12 / 24",
    stat_1_v: "months full-time / part-time",
    stat_2_k: "40",
    stat_2_v: "total credit hours",
    stat_3_k: "9 + 5",
    stat_3_v: "core + elective courses",
    stat_4_k: "MQA",
    stat_4_v: "aligned with NSS & NIMP 2030",
    about_g1_k: "Demand check",
    about_g1_v: "Helps IMEN, UKM test demand for the new M.Sc. Semiconductor Engineering.",
    about_g2_k: "4–6 min",
    about_g2_v: "Quick to finish. No account needed.",
    about_g3_k: "14 course cards",
    about_g3_v: "Skim each card, rate it, then answer a few questions.",
    about_g4_k: "Private",
    about_g4_v: "Answers are used only to improve the programme.",
    about_survey_title: "About the survey",
    about_survey_b1: "Helps IMEN, UKM check demand for the proposed M.Sc. Semiconductor Engineering.",
    about_survey_b2: "Takes about 4–6 minutes. No account needed.",
    about_survey_b3: "You skim 14 short course cards, then answer a few questions.",
    about_survey_b4: "Your answers are used only to improve the programme.",
    about_survey_structure: "Structure",
    about_title: "A programme designed with the industry, for the industry.",
    about_lead:
      "IMEN invites industry leaders, academics, and future students to review the proposed structure — from device physics and cleanroom fabrication to advanced packaging, photonics, and supply chain.",
    cur_eyebrow: "Curriculum Explorer",
    cur_title: "Tap a course to see what you will master.",
    cur_sub:
      "Fourteen courses. One integrated journey from device fundamentals to system-level innovation.",
    cur_core: "Core courses",
    cur_elective: "Elective courses",
    tag_core: "Core",
    tag_elective: "Elective",
    cur_view: "View course",
    cur_rate: "Rate this course",
    cur_close: "Close",
    cur_summary: "Summary",
    cur_overview: "Detailed overview",
    cur_tap: "Tap to explore",
    survey_eyebrow: "Questionnaire",
    survey_title: "Your input shapes the programme.",
    survey_lead:
      "Browse all 14 course cards first, then share your profile and preferences. It takes about 4–6 minutes, and your responses are used solely to evaluate and refine the proposed programme.",
    gate_eyebrow: "Before you begin",
    gate_title: "Your profile first,<br />then three short steps.",
    gate_lead:
      "We collect a few details before the questionnaire so every response can be understood in context. You can edit your profile at any time before submitting.",
    gate_profile: "Your profile",
    gate_priorities: "What matters to you",
    gate_privacy: "Your details are used only to analyse responses to this study.",
    gate_start: "Proceed",
    gate_note: "3 sections · about 4–6 minutes · no account needed",
    drag_hint: "Drag to explore",
    hero_die_label: "Drag me",
    hero_unlock_hint: "Drag the chip onto the glowing socket",
    hero_unlock_done: "Chip seated — circuit online",
    lock_title: "The curriculum is still locked",
    lock_hint: "Place the gold die onto its socket on the wafer above to unlock the curriculum explorer and the survey.",
    lock_cta: "Go place the die",
    read_title: "Explore the 14 courses",
    read_tag: "Read",
    cur_hint: "Tap a card — about 20 seconds each.",
    cur_revisit: "Review again",
    stage_foundations: "Foundations",
    stage_foundations_d: "Physics, fabrication processes and materials science.",
    stage_fabrication: "Fabrication & Integration",
    stage_fabrication_d: "Turning wafers into packaged, testable microelectronic systems.",
    stage_practice: "Research & Professional Practice",
    stage_practice_d: "Methods, management and the two-phase final project.",
    stage_specialise: "Specialisation Electives",
    stage_specialise_d: "Pick the tracks that match your career direction.",
    sheet_learn: "What you will learn",
    sheet_full: "Read the full overview",
    sheet_next: "Got it — next course",
    sheet_prev: "Back",
    sheet_of: "of",
    read_hint: "You are encouraged to skim the summary of each course in the Curriculum Explorer — about 20 seconds per card — before beginning the survey.",
    read_done: "All courses read — you can begin.",
    read_mark_all: "Mark all as read",
    err_read: "Almost there — kindly skim through all 14 course summaries in the Curriculum Explorer before beginning the survey.",
    gate_profile_short: "Profile",
    step_of: "Section {n} of {total}",
    step_1: "Curriculum",
    step_2: "Preferences",
    step_3: "Review",
    h_courses: "Explore the curriculum",
    h_prefs: "Your preferences",
    h_review: "Review & submit",
    courses_lead:
      "Go through the courses one by one — no rating needed. Tap the button on each card to move on.",
    reading_section: "Reading section",
    info_heading: "Programme at a glance",
    info_credits: "Total credits",
    info_duration: "Duration",
    info_months: "12 / 24 months",
    info_full_part: "full-time / part-time",
    info_core_n: "Core courses",
    info_elective_n: "Elective courses",
    info_core_split: "Core · 34 credits",
    info_elective_split: "Elective · 6 credits",
    info_split_note: "85% core · 15% elective",
    info_flow: "How the programme builds",
    info_stage_1: "Design",
    info_stage_2: "Device",
    info_stage_3: "Process",
    info_stage_4: "Characterization",
    info_stage_5: "Packaging",
    info_note: "MQA-aligned · Supports the National Semiconductor Strategy and NIMP 2030",
    read_course: "Read",
    courses_word: "courses",
    core_group: "Core courses",
    elective_group: "Elective courses",
    prefs_lead: "Three quick choices, then an optional comment.",
    p_email: "Email",
    p_email_ph: "you@organisation.com",
    p_name: "Name (Optional)",
    p_name_ph: "Your full name",
    p_category: "Respondent Category / Employment Sector",
    p_select: "Select your category…",
    p_other: "Please specify",
    p_other_ph: "Type your answer",
    p_criteria_lead: "Select all that apply.",
    p_core_lead: "Rate each core course from 1 to 5.",
    p_elective_lead: "Rate each elective course from 1 to 5.",
    p_scale: "Scale",
    p_mode: "Recommended Study Mode Delivery",
    p_fee: "Suggested Reasonable Tuition Fee Rate (Local / Citizen)",
    p_fee_other: "Please specify amount (RM)",
    p_interest: "Interest in the programme",
    p_feedback: "Suggestions for improvement",
    p_feedback_ph:
      "Share any courses, skills, or improvements you would recommend…",
    card_rate_q: "How suitable and important is this course for the programme?",
    err_rate: "Please rate this course (1–5) before moving on.",
    review_rated: "courses rated",
    scale_1: "Highly Unsuitable",
    scale_2: "Unsuitable",
    scale_3: "Moderate",
    scale_4: "Suitable",
    scale_5: "Highly Suitable",
    card_course_label: "Course",
    card_of: "Card {n} of {total}",
    card_prev: "Previous card",
    card_next: "Got it — next course",
    card_continue: "Continue to questions",
    btn_back: "Back",
    btn_next: "Continue",
    btn_submit: "Submit response",
    btn_review: "Review answers",
    review_title: "Almost there",
    review_lead: "Please confirm your details before submitting.",
    review_profile: "Profile",
    review_courses: "Curriculum",
    review_prefs: "Preferences",
    review_feedback: "Feedback",
    review_edit: "Edit",
    thank_title: "Thank you",
    thank_msg:
      "Your response has been recorded. Your insight will directly shape the development of the Master of Science in Semiconductor Engineering programme at IMEN, UKM.",
    thank_again: "Submit another response",
    err_title: "Action needed",
    err_required: "Please complete all required fields before continuing.",
    err_email: "Please enter a valid email address.",
    err_other: "Please specify your answer for “Others”.",
    err_criteria: "Please select at least one criterion.",
    err_courses: "Please rate every course before continuing.",
    err_core: "Please rate every core course.",
    err_elective: "Please rate every elective course.",
    err_prefs: "Please complete your study mode, fee, and interest selection.",
    err_submit: "We could not submit your response. Please try again.",
    lang_note: "Bahasa Melayu interface available",
    demo_badge: "Demo mode — responses are stored locally",
    required: "Required",
    optional: "Optional",
    of: "of",
    courses_rated: "courses rated"
  },
  ms: {
    brand: "IMEN · UKM",
    nav_about: "Tentang",
    nav_curriculum: "Kurikulum",
    nav_survey: "Soal Selidik",
    nav_lang: "EN",
    hero_eyebrow: "Kajian Kebolehpasaran · Program Baharu",
    hero_title_1: "Sarjana Sains dalam",
    hero_title_accent: "Kejuruteraan Semikonduktor",
    hero_lead:
      "Bantu membentuk program yang selari dengan Strategi Semikonduktor Negara. Pandangan anda memastikan kurikulum relevan dengan industri, mematuhi MQA, dan berorientasikan masa depan.",
    hero_cta_start: "Mulakan soal selidik",
    hero_cta_explore: "Terokai kurikulum",
    hero_meta_mode: "Mod pengajian",
    hero_meta_credits: "Jumlah kredit",
    hero_meta_split: "Struktur",
    hero_core_elective: "Teras 34 · Elektif 6",
    hero_scroll: "Tatal untuk teroka",
    stat_1_k: "12 / 24",
    stat_1_v: "bulan sepenuh masa / separuh masa",
    stat_2_k: "40",
    stat_2_v: "jumlah jam kredit",
    stat_3_k: "9 + 5",
    stat_3_v: "kursus teras + elektif",
    stat_4_k: "MQA",
    stat_4_v: "selari dengan NSS & NIMP 2030",
    about_g1_k: "Semakan permintaan",
    about_g1_v: "Membantu IMEN, UKM menilai permintaan untuk M.Sc. Kejuruteraan Semikonduktor yang baharu.",
    about_g2_k: "4–6 minit",
    about_g2_v: "Cepat disiapkan. Tiada akaun diperlukan.",
    about_g3_k: "14 kad kursus",
    about_g3_v: "Semak setiap kad, nilaikan, kemudian jawab beberapa soalan.",
    about_g4_k: "Peribadi",
    about_g4_v: "Jawapan digunakan untuk menambah baik program sahaja.",
    about_survey_title: "Tentang soal selidik",
    about_survey_b1: "Membantu IMEN, UKM menilai permintaan untuk M.Sc. Kejuruteraan Semikonduktor yang dicadangkan.",
    about_survey_b2: "Mengambil masa kira-kira 4–6 minit. Tiada akaun diperlukan.",
    about_survey_b3: "Anda menyemak 14 kad kursus ringkas, kemudian menjawab beberapa soalan.",
    about_survey_b4: "Jawapan anda digunakan untuk menambah baik program sahaja.",
    about_survey_structure: "Struktur",
    about_title: "Program direka bersama industri, untuk industri.",
    about_lead:
      "IMEN menjemput pemimpin industri, ahli akademik, dan bakal pelajar menilai cadangan struktur ini — daripada fizik peranti dan fabrikasi bilik bersih sehingga pembungkusan termaju, fotonik, dan rantaian bekalan.",
    cur_eyebrow: "Penjelajah Kurikulum",
    cur_title: "Ketik kursus untuk melihat apa yang akan anda kuasai.",
    cur_sub:
      "Empat belas kursus. Satu perjalanan bersepadu daripada asas peranti hingga inovasi peringkat sistem.",
    cur_core: "Kursus teras",
    cur_elective: "Kursus elektif",
    tag_core: "Teras",
    tag_elective: "Elektif",
    cur_view: "Lihat kursus",
    cur_rate: "Nilai kursus ini",
    cur_close: "Tutup",
    cur_summary: "Ringkasan",
    cur_overview: "Gambaran terperinci",
    cur_tap: "Ketik untuk teroka",
    survey_eyebrow: "Soal Selidik",
    survey_title: "Input anda membentuk program ini.",
    survey_lead:
      "Semak kesemua 14 kad kursus dahulu, kemudian kongsi profil dan pilihan anda. Ia mengambil masa kira-kira 4–6 minit, dan jawapan anda digunakan semata-mata untuk menilai dan menambah baik cadangan program.",
    gate_eyebrow: "Sebelum anda mula",
    gate_title: "Profil anda dahulu,<br />kemudian tiga langkah ringkas.",
    gate_lead:
      "Kami mengumpul beberapa butiran sebelum soal selidik supaya setiap jawapan dapat difahami dalam konteks. Anda boleh mengubah profil anda pada bila-bila masa sebelum menghantar.",
    gate_profile: "Profil anda",
    gate_priorities: "Apa yang penting bagi anda",
    gate_privacy: "Butiran anda digunakan hanya untuk menganalisis jawapan kajian ini.",
    gate_start: "Teruskan",
    gate_note: "3 bahagian · kira-kira 4–6 minit · tiada akaun diperlukan",
    drag_hint: "Seret untuk teroka",
    hero_die_label: "Seret saya",
    hero_unlock_hint: "Seret cip ke soket bercahaya",
    hero_unlock_done: "Cip dipasang — litar aktif",
    lock_title: "Kurikulum masih berkunci",
    lock_hint: "Letakkan cip emas pada soketnya di atas wafer di bahagian atas untuk membuka peneroka kurikulum dan soal selidik.",
    lock_cta: "Pergi letakkan cip",
    read_title: "Terokai 14 kursus",
    read_tag: "Dibaca",
    cur_hint: "Ketik kad — lebih kurang 20 saat setiap satu.",
    cur_revisit: "Semak semula",
    stage_foundations: "Asas",
    stage_foundations_d: "Fizik, proses fabrikasi dan sains bahan.",
    stage_fabrication: "Fabrikasi & Integrasi",
    stage_fabrication_d: "Mentransformasi wafer kepada sistem mikroelektronik terbungkus yang boleh diuji.",
    stage_practice: "Penyelidikan & Amalan Profesional",
    stage_practice_d: "Kaedah, pengurusan dan projek akhir dua fasa.",
    stage_specialise: "Elektif Pengkhususan",
    stage_specialise_d: "Pilih trek yang sepadan dengan hala tuju kerjaya anda.",
    sheet_learn: "Apa yang akan anda pelajari",
    sheet_full: "Baca gambaran penuh",
    sheet_next: "Faham — kursus seterusnya",
    sheet_prev: "Kembali",
    sheet_of: "daripada",
    read_hint: "Anda digalakkan untuk menyemak ringkasan setiap kursus dalam Peneroka Kurikulum — lebih kurang 20 saat setiap kad — sebelum memulakan soal selidik.",
    read_done: "Semua kursus telah dibaca — anda boleh mula.",
    read_mark_all: "Tanda semua sebagai dibaca",
    err_read: "Hampir selesai — sila semak kesemua 14 ringkasan kursus dalam Peneroka Kurikulum sebelum memulakan soal selidik.",
    gate_profile_short: "Profil",
    step_of: "Bahagian {n} daripada {total}",
    step_1: "Kurikulum",
    step_2: "Keutamaan",
    step_3: "Semakan",
    h_courses: "Terokai kurikulum",
    h_prefs: "Keutamaan anda",
    h_review: "Semak & hantar",
    courses_lead:
      "Semak kursus satu demi satu — tiada penilaian diperlukan. Ketik butang pada setiap kad untuk meneruskan.",
    reading_section: "Bahagian bacaan",
    info_heading: "Program secara ringkas",
    info_credits: "Jumlah kredit",
    info_duration: "Tempoh",
    info_months: "12 / 24 bulan",
    info_full_part: "sepenuh masa / separuh masa",
    info_core_n: "Kursus teras",
    info_elective_n: "Kursus elektif",
    info_core_split: "Teras · 34 kredit",
    info_elective_split: "Elektif · 6 kredit",
    info_split_note: "85% teras · 15% elektif",
    info_flow: "Bagaimana program dibina",
    info_stage_1: "Reka bentuk",
    info_stage_2: "Peranti",
    info_stage_3: "Proses",
    info_stage_4: "Pencirian",
    info_stage_5: "Pembungkusan",
    info_note: "Selari MQA · Menyokong Strategi Semikonduktor Nasional dan NIMP 2030",
    read_course: "Baca",
    courses_word: "kursus",
    core_group: "Kursus teras",
    elective_group: "Kursus elektif",
    prefs_lead: "Tiga pilihan pantas, kemudian komen pilihan.",
    p_email: "E-mel",
    p_email_ph: "anda@organisasi.com",
    p_name: "Nama (Pilihan)",
    p_name_ph: "Nama penuh anda",
    p_category: "Kategori Responden / Sektor Pekerjaan",
    p_select: "Pilih kategori anda…",
    p_other: "Sila nyatakan",
    p_other_ph: "Taip jawapan anda",
    p_criteria_lead: "Pilih semua yang berkenaan.",
    p_core_lead: "Nilai setiap kursus teras dari 1 hingga 5.",
    p_elective_lead: "Nilai setiap kursus elektif dari 1 hingga 5.",
    p_scale: "Skala",
    p_mode: "Cadangan Mod Penyampaian Pengajian",
    p_fee: "Cadangan Kadar Yuran Pengajian Berpatutan (Tempatan / Warganegara)",
    p_fee_other: "Sila nyatakan jumlah (RM)",
    p_interest: "Minat terhadap program",
    p_feedback: "Cadangan penambahbaikan",
    p_feedback_ph:
      "Kongsi kursus, kemahiran, atau penambahbaikan yang anda syorkan…",
    card_rate_q: "Sejauh mana kursus ini sesuai dan penting untuk program ini?",
    err_rate: "Sila nilai kursus ini (1–5) sebelum meneruskan.",
    review_rated: "kursus dinilai",
    scale_1: "Sangat Tidak Sesuai",
    scale_2: "Tidak Sesuai",
    scale_3: "Sederhana",
    scale_4: "Sesuai",
    scale_5: "Sangat Sesuai",
    card_course_label: "Kursus",
    card_of: "Kad {n} daripada {total}",
    card_prev: "Kad sebelumnya",
    card_next: "Faham — kursus seterusnya",
    card_continue: "Teruskan ke soalan",
    btn_back: "Kembali",
    btn_next: "Teruskan",
    btn_submit: "Hantar jawapan",
    btn_review: "Semak jawapan",
    review_title: "Hampir selesai",
    review_lead: "Sila sahkan maklumat anda sebelum menghantar.",
    review_profile: "Profil",
    review_courses: "Kurikulum",
    review_prefs: "Keutamaan",
    review_feedback: "Maklum balas",
    review_edit: "Ubah",
    thank_title: "Terima kasih",
    thank_msg:
      "Jawapan anda telah direkodkan. Pandangan anda akan membentuk pembangunan program Sarjana Sains Kejuruteraan Semikonduktor di IMEN, UKM.",
    thank_again: "Hantar jawapan lain",
    err_title: "Tindakan diperlukan",
    err_required: "Sila lengkapkan semua ruangan wajib sebelum meneruskan.",
    err_email: "Sila masukkan alamat e-mel yang sah.",
    err_other: "Sila nyatakan jawapan anda bagi “Lain-lain”.",
    err_criteria: "Sila pilih sekurang-kurangnya satu kriteria.",
    err_courses: "Sila nilai setiap kursus sebelum meneruskan.",
    err_core: "Sila nilai setiap kursus teras.",
    err_elective: "Sila nilai setiap kursus elektif.",
    err_prefs: "Sila lengkapkan mod pengajian, yuran, dan pilihan minat anda.",
    err_submit: "Kami tidak dapat menghantar jawapan anda. Sila cuba lagi.",
    lang_note: "Antara muka Bahasa Melayu tersedia",
    demo_badge: "Mod demo — jawapan disimpan secara setempat",
    required: "Wajib",
    optional: "Pilihan",
    of: "daripada",
    courses_rated: "kursus dinilai"
  }
};

const t = (lang, key, vars) => {
  let str = (I18N[lang] && I18N[lang][key]) || I18N.en[key] || key;
  if (vars) {
    Object.keys(vars).forEach((k) => {
      str = str.replace(new RegExp("\\{" + k + "\\}", "g"), vars[k]);
    });
  }
  return str;
};
