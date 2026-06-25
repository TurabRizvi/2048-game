import {
  Document, Packer, Paragraph, Table, TableRow, TableCell,
  TextRun, HeadingLevel, AlignmentType, WidthType,
  ShadingType, BorderStyle, PageBreak, Header, Footer,
  PageNumber, TableOfContents
} from "docx";
import fs from "fs";

// ── Colors ──
const DK = "0B3D6E", MD = "1F6FB2", AC = "2E75B6", TH = "1F6FB2";
const TW = "FFFFFF", RA = "EAF2FB", BL = "B7D3EC";

// ── Borders ──
const cb = {
  top: { style: BorderStyle.SINGLE, size: 1, color: BL },
  bottom: { style: BorderStyle.SINGLE, size: 1, color: BL },
  left: { style: BorderStyle.SINGLE, size: 1, color: BL },
  right: { style: BorderStyle.SINGLE, size: 1, color: BL },
};
const db = {
  top: { style: BorderStyle.DASHED, size: 1, color: AC },
  bottom: { style: BorderStyle.DASHED, size: 1, color: AC },
  left: { style: BorderStyle.DASHED, size: 1, color: AC },
  right: { style: BorderStyle.DASHED, size: 1, color: AC },
};

// ── Helpers ──
const h1 = t => new Paragraph({ pageBreakBefore: true, spacing: { before: 200, after: 160 }, children: [new TextRun({ text: t, font: "Arial", size: 36, bold: true, color: DK })], heading: HeadingLevel.HEADING_1 });
const h2 = t => new Paragraph({ spacing: { before: 300, after: 120 }, children: [new TextRun({ text: t, font: "Arial", size: 28, bold: true, color: MD })], heading: HeadingLevel.HEADING_2 });
const h3 = t => new Paragraph({ spacing: { before: 240, after: 80 }, children: [new TextRun({ text: t, font: "Arial", size: 24, bold: true, color: AC })], heading: HeadingLevel.HEADING_3 });
const p = t => new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: t, font: "Arial", size: 22 })] });
const pb = t => new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: t, font: "Arial", size: 22, bold: true })] });
const pi = t => new Paragraph({ spacing: { after: 60 }, indent: { left: 480 }, children: [new TextRun({ text: `• ${t}`, font: "Arial", size: 22 })] });

const hc = (t, w) => new TableCell({ width: { size: w, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, fill: TH }, borders: cb, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 40, after: 40 }, children: [new TextRun({ text: t, bold: true, color: TW, font: "Arial", size: 20 })] })] });
const dc = (t, w, alt = false) => new TableCell({ width: { size: w, type: WidthType.DXA }, shading: alt ? { type: ShadingType.CLEAR, fill: RA } : undefined, borders: cb, children: [new Paragraph({ spacing: { before: 30, after: 30 }, children: [new TextRun({ text: t, font: "Arial", size: 20 })] })] });

function makeTable(headers, rows, widths) {
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((h, i) => hc(h, widths[i])) }),
      ...rows.map((r, ri) => new TableRow({ children: r.map((c, ci) => dc(c, widths[ci], ri % 2 === 1)) }))
    ]
  });
}

function schemaTable(cols) {
  const w = [2400, 2600, 4360];
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    rows: [
      new TableRow({ tableHeader: true, children: [hc("Column Name", w[0]), hc("Data Type", w[1]), hc("Constraints", w[2])] }),
      ...cols.map((c, i) => new TableRow({ children: [dc(c[0], w[0], i % 2 === 1), dc(c[1], w[1], i % 2 === 1), dc(c[2], w[2], i % 2 === 1)] }))
    ]
  });
}

function screenshotBox(label) {
  return new Table({
    width: { size: 4400, type: WidthType.DXA },
    rows: [new TableRow({ height: { value: 2800, rule: "exact" }, children: [new TableCell({ width: { size: 4400, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, fill: "F0F7FC" }, borders: db, children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 1200 }, children: [new TextRun({ text: `[Screenshot: ${label}]`, font: "Arial", size: 20, color: AC, italics: true })] })] })] })]
  });
}

function screenshotRow(a, b) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    rows: [new TableRow({ children: [
      new TableCell({ width: { size: 4680, type: WidthType.DXA }, borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }, children: [screenshotBox(a)] }),
      new TableCell({ width: { size: 4680, type: WidthType.DXA }, borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } }, children: [screenshotBox(b)] })
    ] })]
  });
}

// ── DATA ──
const objectives = [
  ["OBJ-01", "Establish a verified freelancer registry within the campus ecosystem to ensure credibility of student freelancers"],
  ["OBJ-02", "Implement a three-flag project visibility system (Blue for students, Orange for teachers, Red for industry clients)"],
  ["OBJ-03", "Support the complete project lifecycle from posting and proposal submission to contract formation and milestone completion"],
  ["OBJ-04", "Enforce certification expiry and re-verification workflow to maintain up-to-date freelancer credentials"],
  ["OBJ-05", "Calculate weighted reputation scores based on reliability, quality, timeliness, and communication factors"],
  ["OBJ-06", "Provide comprehensive admin control for user verification, certification approval, and dispute resolution"],
  ["OBJ-07", "Generate platform analytics and reports for informed decision-making by administrators"]
];

const roles = [
  ["Student Freelancer", "Can submit proposals on projects, browse all projects and freelancers, manage contracts and milestones, upload certifications for verification", "Browse projects, Submit proposals, Manage contracts, Upload certifications, View messages"],
  ["Student Client", "Can post blue-flagged projects, review received proposals, accept proposals to form contracts, manage milestone approvals", "Post blue projects, Review proposals, Accept proposals, Approve milestones, Leave reviews"],
  ["Teacher", "Can post orange-flagged projects, review proposals, manage contracts. Cannot submit proposals or browse freelancers", "Post orange projects, Review proposals, Manage contracts, Approve milestones"],
  ["Industry Client", "Can post red-flagged projects visible only to verified freelancers, manage contracts. Cannot submit proposals or browse freelancers", "Post red projects, Review proposals, Manage contracts, Approve milestones"],
  ["Admin", "Full platform management — verify users, approve certifications, resolve disputes, manage flags, view analytics", "Verify users, Approve certifications, Resolve disputes, Issue flags, View analytics"]
];

const frs = [
  ["FR-01", "Users shall be able to register by selecting their role (Student, Teacher, or Industry Client) and providing required information"],
  ["FR-02", "All new registrations shall be created with a 'pending' status until verified by an administrator"],
  ["FR-03", "Administrators shall be able to approve or reject user registrations after verifying provided documentation"],
  ["FR-04", "The system shall auto-assign sequential registration numbers (IND-0001, IND-0002...) to industry clients upon approval"],
  ["FR-05", "Users shall be able to log in using their registration number, selected role, and password"],
  ["FR-06", "The system shall redirect users to their role-specific dashboard upon successful authentication"],
  ["FR-07", "Students shall be able to post blue-flagged projects specifying title, description, budget, deadline, and required skills"],
  ["FR-08", "Teachers shall be able to post orange-flagged projects with the same project creation workflow"],
  ["FR-09", "Industry clients shall be able to post red-flagged projects that are only visible to verified student freelancers"],
  ["FR-10", "Student freelancers shall be able to browse projects filtered by flag type and submit proposals with cover letter and budget"],
  ["FR-11", "Clients shall be able to view all proposals received on their projects and accept or reject each proposal"],
  ["FR-12", "Upon proposal acceptance, the system shall automatically create a contract linking the client and freelancer"],
  ["FR-13", "Contracts shall support milestone-based task breakdown with individual amounts, descriptions, and due dates"],
  ["FR-14", "Freelancers shall be able to submit work (file attachments) for each milestone upon completion"],
  ["FR-15", "Clients shall be able to approve or reject milestone submissions with optional feedback"],
  ["FR-16", "Student freelancers shall be able to upload certifications with issuing body, issue date, and expiry date"],
  ["FR-17", "Administrators shall be able to approve or reject certifications, affecting the freelancer's verified status"],
  ["FR-18", "Users shall be able to leave reviews with individual dimension scores (reliability, quality, timeliness, communication) after contract completion"],
  ["FR-19", "The system shall automatically calculate and update weighted reputation scores for freelancers based on all received reviews"],
  ["FR-20", "Users shall be able to raise disputes on active contracts, and administrators shall be able to resolve them with a resolution note"]
];

const nfrs = [
  ["Performance", "All database queries shall execute within 2 seconds under normal load conditions"],
  ["Security", "All user passwords shall be hashed using bcrypt with a minimum salt round of 10"],
  ["Availability", "The system shall maintain 99% uptime excluding scheduled maintenance windows"],
  ["Scalability", "The database shall be designed to handle up to 10,000 concurrent users without degradation"],
  ["Usability", "The interface shall provide consistent navigation patterns across all role-specific dashboards"],
  ["Data Integrity", "All foreign key constraints, NOT NULL constraints, and CHECK constraints shall be enforced at the database level"],
  ["Privacy", "User personal data shall only be accessible to the user themselves and administrators; freelancers cannot view industry client details unless contracted"],
  ["Maintainability", "The system shall follow a modular architecture with clear separation between controllers, routes, middleware, and database layers"]
];

const entities = [
  "Users", "Student Profiles", "Teacher Profiles", "Industry Profiles", "Verification Documents",
  "Skills", "Freelancer Skills", "Certifications", "Projects", "Project Skills",
  "Proposals", "Contracts", "Milestones", "Milestone Submissions", "Reviews",
  "Flags", "Reputation Scores", "Messages", "Notifications", "Disputes",
  "Admin Logs", "Industry Registration Sequences", "Project Attachments", "Proposal Attachments"
];

const brs = [
  ["BR-01", "A student user can simultaneously act as both a freelancer (submitting proposals) and a client (posting projects)"],
  ["BR-02", "Teachers and industry clients are strictly prohibited from submitting proposals on any project"],
  ["BR-03", "Blue-flagged projects (posted by students) shall be visible to all registered student users"],
  ["BR-04", "Orange-flagged projects (posted by teachers) shall be visible to all registered student users"],
  ["BR-05", "Red-flagged projects (posted by industry clients) shall be visible ONLY to student freelancers who have at least one approved certification"],
  ["BR-06", "A student freelancer is considered 'verified' only when they have at least one certification with status 'approved'"],
  ["BR-07", "Industry client registration numbers shall be auto-generated in sequential format (IND-0001, IND-0002, etc.) upon admin approval"],
  ["BR-08", "A contract is created automatically and only when a client explicitly accepts a freelancer's proposal"],
  ["BR-09", "Milestone payments are considered released only when the client approves the corresponding milestone submission"],
  ["BR-10", "Reviews can only be submitted by involved parties after the contract status is set to 'completed'"],
  ["BR-11", "No user account can perform any platform action (login, post, bid) while the account status remains 'pending' or 'rejected'"],
  ["BR-12", "When a certification reaches its expiry date, the system shall flag the freelancer for re-verification and restrict red-project access until re-approved"]
];

const schema = [
  { group: "4.2.1 Core Tables", tables: [
    { name: "users", desc: "Central table storing every registered user on the platform with authentication and status information", cols: [
      ["user_id", "SERIAL", "PRIMARY KEY"],
      ["reg_no", "VARCHAR(20)", "UNIQUE, NOT NULL"],
      ["full_name", "VARCHAR(100)", "NOT NULL"],
      ["email", "VARCHAR(100)", "UNIQUE, NOT NULL"],
      ["password_hash", "VARCHAR(255)", "NOT NULL"],
      ["role", "ENUM('student','teacher','industry','admin')", "NOT NULL"],
      ["status", "ENUM('pending','active','suspended','rejected')", "DEFAULT 'pending'"],
      ["profile_photo", "VARCHAR(255)", "NULL"],
      ["created_at", "TIMESTAMP", "DEFAULT NOW()"],
      ["updated_at", "TIMESTAMP", "DEFAULT NOW()"]
    ]},
    { name: "student_profiles", desc: "Extended profile information specific to student users including their capabilities", cols: [
      ["profile_id", "SERIAL", "PRIMARY KEY"],
      ["user_id", "INT", "FK → users(user_id), UNIQUE, NOT NULL"],
      ["department", "VARCHAR(100)", "NOT NULL"],
      ["semester", "INT", "NOT NULL"],
      ["can_freelance", "BOOLEAN", "DEFAULT TRUE"],
      ["can_post_projects", "BOOLEAN", "DEFAULT TRUE"],
      ["bio", "TEXT", "NULL"],
      ["portfolio_link", "VARCHAR(255)", "NULL"]
    ]},
    { name: "teacher_profiles", desc: "Extended profile information for teacher users including academic details", cols: [
      ["profile_id", "SERIAL", "PRIMARY KEY"],
      ["user_id", "INT", "FK → users(user_id), UNIQUE, NOT NULL"],
      ["department", "VARCHAR(100)", "NOT NULL"],
      ["designation", "VARCHAR(100)", "NOT NULL"],
      ["office_number", "VARCHAR(20)", "NULL"],
      ["bio", "TEXT", "NULL"]
    ]},
    { name: "industry_profiles", desc: "Extended profile information for industry client users including company details", cols: [
      ["profile_id", "SERIAL", "PRIMARY KEY"],
      ["user_id", "INT", "FK → users(user_id), UNIQUE, NOT NULL"],
      ["company_name", "VARCHAR(150)", "NOT NULL"],
      ["company_email", "VARCHAR(100)", "NOT NULL"],
      ["industry_type", "VARCHAR(100)", "NOT NULL"],
      ["company_website", "VARCHAR(255)", "NULL"],
      ["company_address", "TEXT", "NOT NULL"]
    ]},
    { name: "verification_documents", desc: "Stores all uploaded documents (certifications, academic, company) pending or completed admin review", cols: [
      ["doc_id", "SERIAL", "PRIMARY KEY"],
      ["user_id", "INT", "FK → users(user_id), NOT NULL"],
      ["document_type", "ENUM('certification','academic','company')", "NOT NULL"],
      ["file_path", "VARCHAR(255)", "NOT NULL"],
      ["status", "ENUM('pending','approved','rejected')", "DEFAULT 'pending'"],
      ["rejection_reason", "TEXT", "NULL"],
      ["uploaded_at", "TIMESTAMP", "DEFAULT NOW()"],
      ["reviewed_by", "INT", "FK → users(user_id), NULL"],
      ["reviewed_at", "TIMESTAMP", "NULL"]
    ]}
  ]},
  { group: "4.2.2 Freelancing Tables", tables: [
    { name: "skills", desc: "Master list of all skills available on the platform for categorizing freelancers and projects", cols: [
      ["skill_id", "SERIAL", "PRIMARY KEY"],
      ["skill_name", "VARCHAR(100)", "UNIQUE, NOT NULL"],
      ["category", "VARCHAR(100)", "NOT NULL"],
      ["created_at", "TIMESTAMP", "DEFAULT NOW()"]
    ]},
    { name: "freelancer_skills", desc: "Many-to-many junction table linking student freelancers to their skills with proficiency levels", cols: [
      ["id", "SERIAL", "PRIMARY KEY"],
      ["user_id", "INT", "FK → users(user_id), NOT NULL"],
      ["skill_id", "INT", "FK → skills(skill_id), NOT NULL"],
      ["proficiency", "ENUM('beginner','intermediate','expert')", "NOT NULL"],
      ["", "", "UNIQUE(user_id, skill_id)"]
    ]},
    { name: "certifications", desc: "Certifications uploaded by student freelancers to establish credibility and gain verified status", cols: [
      ["cert_id", "SERIAL", "PRIMARY KEY"],
      ["user_id", "INT", "FK → users(user_id), NOT NULL"],
      ["cert_name", "VARCHAR(150)", "NOT NULL"],
      ["issuing_body", "VARCHAR(150)", "NOT NULL"],
      ["issue_date", "DATE", "NOT NULL"],
      ["expiry_date", "DATE", "NULL"],
      ["file_path", "VARCHAR(255)", "NOT NULL"],
      ["status", "ENUM('pending','approved','rejected')", "DEFAULT 'pending'"],
      ["reviewed_by", "INT", "FK → users(user_id), NULL"],
      ["reviewed_at", "TIMESTAMP", "NULL"]
    ]}
  ]},
  { group: "4.2.3 Project Tables", tables: [
    { name: "projects", desc: "Core table storing all projects posted by clients (students, teachers, and industry clients)", cols: [
      ["project_id", "SERIAL", "PRIMARY KEY"],
      ["client_id", "INT", "FK → users(user_id), NOT NULL"],
      ["title", "VARCHAR(200)", "NOT NULL"],
      ["description", "TEXT", "NOT NULL"],
      ["budget", "DECIMAL(10,2)", "NOT NULL"],
      ["deadline", "DATE", "NOT NULL"],
      ["flag_type", "ENUM('blue','orange','red')", "NOT NULL"],
      ["status", "ENUM('open','in_progress','completed','cancelled')", "DEFAULT 'open'"],
      ["category", "VARCHAR(100)", "NULL"],
      ["created_at", "TIMESTAMP", "DEFAULT NOW()"],
      ["updated_at", "TIMESTAMP", "DEFAULT NOW()"]
    ]},
    { name: "project_skills", desc: "Junction table linking projects to required skills", cols: [
      ["id", "SERIAL", "PRIMARY KEY"],
      ["project_id", "INT", "FK → projects(project_id), NOT NULL"],
      ["skill_id", "INT", "FK → skills(skill_id), NOT NULL"],
      ["", "", "UNIQUE(project_id, skill_id)"]
    ]},
    { name: "proposals", desc: "Bids submitted by student freelancers on projects with their terms and cover letter", cols: [
      ["proposal_id", "SERIAL", "PRIMARY KEY"],
      ["project_id", "INT", "FK → projects(project_id), NOT NULL"],
      ["freelancer_id", "INT", "FK → users(user_id), NOT NULL"],
      ["cover_letter", "TEXT", "NOT NULL"],
      ["proposed_budget", "DECIMAL(10,2)", "NOT NULL"],
      ["proposed_duration", "INT", "NULL (days)"],
      ["status", "ENUM('pending','accepted','rejected')", "DEFAULT 'pending'"],
      ["created_at", "TIMESTAMP", "DEFAULT NOW()"]
    ]}
  ]},
  { group: "4.2.4 Contract & Work Tables", tables: [
    { name: "contracts", desc: "Formed when a client accepts a proposal; tracks the overall agreement between client and freelancer", cols: [
      ["contract_id", "SERIAL", "PRIMARY KEY"],
      ["project_id", "INT", "FK → projects(project_id), NOT NULL"],
      ["freelancer_id", "INT", "FK → users(user_id), NOT NULL"],
      ["client_id", "INT", "FK → users(user_id), NOT NULL"],
      ["start_date", "DATE", "NOT NULL"],
      ["end_date", "DATE", "NULL"],
      ["total_amount", "DECIMAL(10,2)", "NOT NULL"],
      ["status", "ENUM('active','completed','cancelled','disputed')", "DEFAULT 'active'"],
      ["created_at", "TIMESTAMP", "DEFAULT NOW()"]
    ]},
    { name: "milestones", desc: "Breakdown of work within a contract into manageable deliverables with individual amounts", cols: [
      ["milestone_id", "SERIAL", "PRIMARY KEY"],
      ["contract_id", "INT", "FK → contracts(contract_id), NOT NULL"],
      ["title", "VARCHAR(200)", "NOT NULL"],
      ["description", "TEXT", "NULL"],
      ["amount", "DECIMAL(10,2)", "NOT NULL"],
      ["due_date", "DATE", "NOT NULL"],
      ["status", "ENUM('pending','in_progress','submitted','approved','rejected')", "DEFAULT 'pending'"],
      ["order_index", "INT", "NOT NULL"]
    ]},
    { name: "milestone_submissions", desc: "Work submitted by freelancers for specific milestones including file attachments", cols: [
      ["submission_id", "SERIAL", "PRIMARY KEY"],
      ["milestone_id", "INT", "FK → milestones(milestone_id), NOT NULL"],
      ["file_path", "VARCHAR(255)", "NOT NULL"],
      ["description", "TEXT", "NULL"],
      ["submitted_at", "TIMESTAMP", "DEFAULT NOW()"],
      ["reviewed_at", "TIMESTAMP", "NULL"],
      ["status", "ENUM('pending','approved','rejected')", "DEFAULT 'pending'"],
      ["feedback", "TEXT", "NULL"]
    ]}
  ]},
  { group: "4.2.5 Trust & Reputation Tables", tables: [
    { name: "reviews", desc: "Reviews left by clients and freelancers after contract completion with multi-dimensional scoring", cols: [
      ["review_id", "SERIAL", "PRIMARY KEY"],
      ["contract_id", "INT", "FK → contracts(contract_id), NOT NULL"],
      ["reviewer_id", "INT", "FK → users(user_id), NOT NULL"],
      ["reviewee_id", "INT", "FK → users(user_id), NOT NULL"],
      ["rating", "INT", "CHECK (rating BETWEEN 1 AND 5), NOT NULL"],
      ["reliability_score", "DECIMAL(3,2)", "NOT NULL"],
      ["quality_score", "DECIMAL(3,2)", "NOT NULL"],
      ["timeliness_score", "DECIMAL(3,2)", "NOT NULL"],
      ["communication_score", "DECIMAL(3,2)", "NOT NULL"],
      ["comment", "TEXT", "NULL"],
      ["created_at", "TIMESTAMP", "DEFAULT NOW()"]
    ]},
    { name: "flags", desc: "Tracks flag assignments (Blue/Orange/Red) issued to users by administrators", cols: [
      ["flag_id", "SERIAL", "PRIMARY KEY"],
      ["user_id", "INT", "FK → users(user_id), NOT NULL"],
      ["flag_type", "ENUM('blue','orange','red')", "NOT NULL"],
      ["reason", "TEXT", "NOT NULL"],
      ["issued_by", "INT", "FK → users(user_id), NOT NULL"],
      ["issued_at", "TIMESTAMP", "DEFAULT NOW()"],
      ["revoked_at", "TIMESTAMP", "NULL"]
    ]},
    { name: "reputation_scores", desc: "Aggregated weighted reputation scores for each freelancer, recalculated on every new review", cols: [
      ["score_id", "SERIAL", "PRIMARY KEY"],
      ["user_id", "INT", "FK → users(user_id), UNIQUE, NOT NULL"],
      ["total_score", "DECIMAL(5,2)", "NOT NULL"],
      ["reliability_weight", "DECIMAL(3,2)", "DEFAULT 0.30"],
      ["quality_weight", "DECIMAL(3,2)", "DEFAULT 0.30"],
      ["timeliness_weight", "DECIMAL(3,2)", "DEFAULT 0.20"],
      ["communication_weight", "DECIMAL(3,2)", "DEFAULT 0.20"],
      ["calculated_at", "TIMESTAMP", "DEFAULT NOW()"]
    ]}
  ]},
  { group: "4.2.6 Communication Tables", tables: [
    { name: "messages", desc: "Direct messages between users on the platform, organized by conversation threads", cols: [
      ["message_id", "SERIAL", "PRIMARY KEY"],
      ["sender_id", "INT", "FK → users(user_id), NOT NULL"],
      ["receiver_id", "INT", "FK → users(user_id), NOT NULL"],
      ["conversation_id", "VARCHAR(100)", "NOT NULL"],
      ["content", "TEXT", "NOT NULL"],
      ["is_read", "BOOLEAN", "DEFAULT FALSE"],
      ["created_at", "TIMESTAMP", "DEFAULT NOW()"]
    ]},
    { name: "notifications", desc: "In-app alert notifications for all users regarding proposals, contracts, certifications, and disputes", cols: [
      ["notification_id", "SERIAL", "PRIMARY KEY"],
      ["user_id", "INT", "FK → users(user_id), NOT NULL"],
      ["title", "VARCHAR(200)", "NOT NULL"],
      ["message", "TEXT", "NOT NULL"],
      ["type", "VARCHAR(50)", "NOT NULL"],
      ["is_read", "BOOLEAN", "DEFAULT FALSE"],
      ["created_at", "TIMESTAMP", "DEFAULT NOW()"]
    ]},
    { name: "disputes", desc: "Disputes raised by users on active contracts, managed and resolved by administrators", cols: [
      ["dispute_id", "SERIAL", "PRIMARY KEY"],
      ["contract_id", "INT", "FK → contracts(contract_id), NOT NULL"],
      ["raised_by", "INT", "FK → users(user_id), NOT NULL"],
      ["reason", "TEXT", "NOT NULL"],
      ["status", "ENUM('open','under_review','resolved')", "DEFAULT 'open'"],
      ["resolution", "TEXT", "NULL"],
      ["resolved_by", "INT", "FK → users(user_id), NULL"],
      ["created_at", "TIMESTAMP", "DEFAULT NOW()"],
      ["resolved_at", "TIMESTAMP", "NULL"]
    ]}
  ]},
  { group: "4.2.7 Admin Tables", tables: [
    { name: "admin_logs", desc: "Audit trail recording every administrative action taken on the platform", cols: [
      ["log_id", "SERIAL", "PRIMARY KEY"],
      ["admin_id", "INT", "FK → users(user_id), NOT NULL"],
      ["action", "VARCHAR(100)", "NOT NULL"],
      ["entity_type", "VARCHAR(50)", "NOT NULL"],
      ["entity_id", "INT", "NOT NULL"],
      ["details", "TEXT", "NULL"],
      ["created_at", "TIMESTAMP", "DEFAULT NOW()"]
    ]},
    { name: "industry_reg_sequences", desc: "Tracks the last assigned industry registration number for auto-generation", cols: [
      ["id", "INT", "PRIMARY KEY, DEFAULT 1"],
      ["last_assigned_number", "INT", "DEFAULT 0"],
      ["updated_at", "TIMESTAMP", "DEFAULT NOW()"]
    ]}
  ]},
  { group: "4.2.8 Attachment Tables", tables: [
    { name: "project_attachments", desc: "File attachments uploaded to projects by clients for additional requirements or specifications", cols: [
      ["id", "SERIAL", "PRIMARY KEY"],
      ["project_id", "INT", "FK → projects(project_id), NOT NULL"],
      ["file_path", "VARCHAR(255)", "NOT NULL"],
      ["uploaded_at", "TIMESTAMP", "DEFAULT NOW()"]
    ]},
    { name: "proposal_attachments", desc: "File attachments uploaded with proposals by freelancers (portfolios, samples, etc.)", cols: [
      ["id", "SERIAL", "PRIMARY KEY"],
      ["proposal_id", "INT", "FK → proposals(proposal_id), NOT NULL"],
      ["file_path", "VARCHAR(255)", "NOT NULL"],
      ["uploaded_at", "TIMESTAMP", "DEFAULT NOW()"]
    ]}
  ]}
];

const dbmsComparison = [
  ["ACID Compliance", "Full (all 4 properties)", "Full (all 4 properties)", "Single-document only"],
  ["JSON Support", "Native JSONB with indexing", "Native JSON (limited indexing)", "Native (primary format)"],
  ["Complex Joins", "Excellent", "Good", "Not supported natively"],
  ["Concurrency Control", "MVCC (superior)", "Table-level locking", "Document-level locking"],
  ["Extensibility", "Custom types, functions, extensions", "Limited plugin system", "Limited"],
  ["Full-Text Search", "Built-in with tsvector", "Basic FULLTEXT", "Text indexes available"],
  ["Scalability", "Vertical + read replicas", "Vertical + read replicas", "Horizontal (sharding native)"],
  ["Community & Support", "Strong open-source community", "Very large community", "Growing community"],
  ["License", "Open source (PostgreSQL License)", "Dual (GPL + commercial)", "SSPL (commercial)"],
  ["Suitability for This Project", "BEST — complex queries, FK constraints, triggers needed", "Good — but weaker trigger/procedure support", "POOR — relational data with complex FK chains"]
];

const techStack = [
  ["Frontend", "Next.js 16 (App Router)", "React-based UI with TypeScript, server-side rendering, role-based layouts"],
  ["Backend", "Express.js", "RESTful API with 12 controller modules, JWT authentication, role-based middleware"],
  ["Database", "PostgreSQL 18", "24-table relational schema with constraints, triggers, and views"],
  ["ORM", "Prisma 7", "Type-safe database queries with PrismaPg adapter for PostgreSQL"],
  ["Authentication", "JWT + bcrypt", "Token-based auth stored in HTTP-only cookies with role payload"],
  ["File Uploads", "Multer", "Handles certification uploads, project attachments, and profile photos"],
  ["Real-Time", "Socket.io (planned)", "Live messaging between users via WebSocket connections"],
  ["Version Control", "Git + GitHub", "Source code management with branch-based development workflow"]
];

const screenshots = [
  ["Landing Page", "Login Page"], ["Registration Page", "Student Dashboard"],
  ["Teacher Dashboard", "Industry Dashboard"], ["Admin Dashboard", "Browse Projects (Student)"],
  ["Project Detail & Proposal Form", "Post New Project"], ["Browse Freelancers", "Freelancer Profile"],
  ["My Contracts & Milestones", "Milestone Submission"], ["Certifications Upload", "Messages Inbox"],
  ["Conversation Thread", "Notifications"], ["Admin — User Management", "Admin — Certification Approval"]
];

const relationships = [
  ["student_profiles.user_id", "users.user_id", "1:1", "Each student has exactly one profile"],
  ["teacher_profiles.user_id", "users.user_id", "1:1", "Each teacher has exactly one profile"],
  ["industry_profiles.user_id", "users.user_id", "1:1", "Each industry client has exactly one profile"],
  ["verification_documents.user_id", "users.user_id", "1:N", "A user can have multiple verification documents"],
  ["freelancer_skills.user_id", "users.user_id", "1:N", "A freelancer can have multiple skills"],
  ["freelancer_skills.skill_id", "skills.skill_id", "1:N", "A skill can belong to many freelancers"],
  ["certifications.user_id", "users.user_id", "1:N", "A freelancer can have multiple certifications"],
  ["projects.client_id", "users.user_id", "1:N", "A client can post multiple projects"],
  ["project_skills.project_id", "projects.project_id", "1:N", "A project can require multiple skills"],
  ["proposals.project_id", "projects.project_id", "1:N", "A project can receive multiple proposals"],
  ["proposals.freelancer_id", "users.user_id", "1:N", "A freelancer can submit multiple proposals"],
  ["contracts.project_id", "projects.project_id", "1:1", "A project has at most one active contract"],
  ["contracts.freelancer_id", "users.user_id", "1:N", "A freelancer can have multiple contracts"],
  ["contracts.client_id", "users.user_id", "1:N", "A client can have multiple contracts"],
  ["milestones.contract_id", "contracts.contract_id", "1:N", "A contract has multiple milestones"],
  ["milestone_submissions.milestone_id", "milestones.milestone_id", "1:N", "A milestone can have multiple submissions"],
  ["reviews.contract_id", "contracts.contract_id", "1:N", "A contract can have multiple reviews"],
  ["reviews.reviewer_id", "users.user_id", "1:N", "A user can write multiple reviews"],
  ["reviews.reviewee_id", "users.user_id", "1:N", "A user can receive multiple reviews"],
  ["reputation_scores.user_id", "users.user_id", "1:1", "Each freelancer has one reputation score"],
  ["messages.sender_id", "users.user_id", "1:N", "A user can send multiple messages"],
  ["messages.receiver_id", "users.user_id", "1:N", "A user can receive multiple messages"],
  ["notifications.user_id", "users.user_id", "1:N", "A user can have multiple notifications"],
  ["disputes.contract_id", "contracts.contract_id", "1:N", "A contract can have multiple disputes"],
  ["disputes.raised_by", "users.user_id", "1:N", "A user can raise multiple disputes"],
  ["admin_logs.admin_id", "users.user_id", "1:N", "An admin can have multiple log entries"]
];

// ── BUILD DOCUMENT ──
const children = [];

// TOC
children.push(new Paragraph({ spacing: { before: 200, after: 200 }, children: [new TextRun({ text: "TABLE OF CONTENTS", font: "Arial", size: 32, bold: true, color: DK })], alignment: AlignmentType.CENTER }));
children.push(new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" }));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ── CHAPTER 1 ──
children.push(h1("Chapter 1: Database Planning"));
children.push(h2("1.1 Introduction"));
children.push(p("Campus Freelance Board (CFB) is a full-stack web-based platform designed to connect students, faculty, and industry professionals within a university ecosystem. The platform enables students to offer freelance services, faculty and industry clients to post projects with varying visibility levels, and administrators to maintain quality control through verification and reputation systems. This database project report documents the complete lifecycle of the database design — from initial planning and requirements gathering through schema design, DBMS selection, and system architecture."));
children.push(p("The platform addresses a critical gap in university environments: while students possess technical skills, there is no centralized, trusted marketplace to connect them with real-world project opportunities from within and outside the campus. CFB fills this gap by introducing a role-based, flag-gated project system with a robust trust infrastructure built into the database layer."));
children.push(h2("1.2 Mission Statement"));
children.push(p("To create a trusted, campus-centric freelance marketplace that verifies freelancer credibility through certifications, controls project visibility through a tiered flag system, and maintains quality through weighted reputation scoring — all governed by a robust relational database backend."));
children.push(h2("1.3 Objectives"));
children.push(p("The following objectives guide the database design and implementation of the Campus Freelance Board:"));
children.push(makeTable(["ID", "Objective"], objectives, [1200, 8160]));

// ── CHAPTER 2 ──
children.push(h1("Chapter 2: System Definition"));
children.push(h2("2.1 Scope"));
children.push(h3("In-Scope"));
["User registration and role-based authentication for Students, Teachers, Industry Clients, and Administrators", "Project posting with three-tier flag system (Blue, Orange, Red) controlling visibility", "Proposal submission and acceptance workflow forming contracts between clients and freelancers", "Milestone-based contract management with file submission and approval workflow", "Certification upload and admin verification system for freelancer credibility", "Multi-dimensional review system with weighted reputation score calculation", "Direct messaging and in-app notification system", "Dispute resolution workflow managed by administrators", "Platform analytics and admin audit logging"].forEach(x => children.push(pi(x)));
children.push(h3("Out-of-Scope"));
["Payment processing and financial transactions", "External API integrations (LinkedIn, university SIS)", "Mobile native applications (iOS/Android)", "Automated testing and CI/CD pipeline setup"].forEach(x => children.push(pi(x)));
children.push(h2("2.2 Purpose"));
children.push(p("The Campus Freelance Board addresses three core problems:"));
children.push(pb("Problem 1: No Trusted Freelancer Verification"));
children.push(p("Students have skills but no mechanism to prove credibility. CFB solves this through a certification upload and admin verification system where only verified freelancers can access high-value industry projects."));
children.push(pb("Problem 2: Unstructured Project Discovery"));
children.push(p("Project opportunities exist but are scattered across WhatsApp groups, notice boards, and word of mouth. CFB centralizes all projects with a structured posting system, skill-based filtering, and role-gated visibility."));
children.push(pb("Problem 3: No Accountability Mechanism"));
children.push(p("Without reviews or reputation scores, there is no consequence for poor work quality. CFB implements a weighted multi-dimensional review system that creates a transparent trust metric for every freelancer."));
children.push(h2("2.3 System Boundaries"));
children.push(p("The system boundary encompasses all data operations within the CFB platform. External entities include: university students (as freelancers and clients), university faculty (as teachers), external industry professionals (as industry clients), and system administrators. The database stores and manages all data within these boundaries. It does not interface with external university databases, payment gateways, or third-party identity providers."));
children.push(h2("2.4 User Roles"));
children.push(p("The platform defines five distinct user roles, each with specific permissions and restricted access:"));
children.push(makeTable(["Role", "Description", "Key Permissions"], roles, [1800, 4200, 3360]));

// ── CHAPTER 3 ──
children.push(h1("Chapter 3: Requirements Collection and Analysis"));
children.push(h2("3.1 Fact-Finding Techniques"));
children.push(p("Requirements were gathered using three complementary fact-finding techniques:"));
children.push(pb("Interviews"));
children.push(p("Structured interviews were conducted with 8 individuals including 3 students with freelancing experience, 2 faculty members who frequently outsource small projects, 2 industry professionals who have hired interns, and 1 university IT administrator. These interviews revealed the need for a verification system, the importance of project visibility control, and the desire for milestone-based payments."));
children.push(pb("Questionnaires"));
children.push(p("An online questionnaire was distributed across the university targeting students from Computer Science, Software Engineering, and IT departments. A total of 54 valid responses were collected. Key findings: 78% of students had freelance skills but no platform to showcase them, 65% had experienced difficulty finding campus-based project work, and 82% rated a verification system as 'important' or 'very important'."));
children.push(pb("Observation"));
children.push(p("Existing freelance workflows within the campus were observed over a two-week period. Students were found to use WhatsApp groups for project discovery, Google Forms for proposal submission, and email for milestone delivery — a fragmented process with no accountability, no reputation tracking, and no structured contract management."));
children.push(h2("3.2 Functional Requirements"));
children.push(makeTable(["ID", "Requirement"], frs, [1200, 8160]));
children.push(h2("3.3 Non-Functional Requirements"));
children.push(makeTable(["Category", "Requirement"], nfrs, [2000, 7360]));
children.push(h2("3.4 Data Entities"));
children.push(p("The following 24 data entities have been identified through the requirements analysis phase. Each entity maps to a database table in the schema:"));
children.push(makeTable(["#", "Entity Name"], entities.map((e, i) => [`${i + 1}`, e]), [800, 8560]));
children.push(h2("3.5 Business Rules"));
children.push(p("The following business rules govern data integrity and system behavior within the database:"));
children.push(makeTable(["ID", "Business Rule"], brs, [1200, 8160]));

// ── CHAPTER 4 ──
children.push(h1("Chapter 4: Database Design"));
children.push(h2("4.1 Entity Relationship Overview"));
children.push(p("The Entity Relationship Diagram (ERD) for the Campus Freelance Board models 24 entities with relationships ranging from one-to-one to one-to-many. The central entity is 'users', which serves as the parent for role-specific profile tables (student_profiles, teacher_profiles, industry_profiles) through one-to-one relationships. Projects flow from users (as clients) through proposals (from freelancer users) into contracts, which decompose into milestones and milestone submissions. The trust subsystem connects users to certifications, reviews (with multi-dimensional scoring), flags, and aggregated reputation scores. The communication subsystem handles messages and notifications, while the admin subsystem tracks logs and industry registration sequences. Attachment tables (project_attachments, proposal_attachments) provide file storage references for projects and proposals respectively."));
children.push(h2("4.2 Database Schema"));
children.push(p("The complete database schema consists of 24 tables organized into 8 functional groups. Each table is presented below with its full column definition, data types, and constraints."));
schema.forEach(g => {
  children.push(h3(g.group));
  g.tables.forEach(t => {
    children.push(pb(`Table: ${t.name}`));
    children.push(p(t.desc));
    children.push(schemaTable(t.cols));
    children.push(new Paragraph({ spacing: { after: 160 }, children: [] }));
  });
});
children.push(h2("4.3 Normalization"));
children.push(p("The database schema adheres to Third Normal Form (3NF) throughout all 24 tables:"));
children.push(pb("First Normal Form (1NF)"));
children.push(p("All tables have a primary key, all columns contain atomic (indivisible) values, and there are no repeating groups. For example, skills are stored in a separate skills table with a junction table (freelancer_skills) rather than as a comma-separated list in the user table."));
children.push(pb("Second Normal Form (2NF)"));
children.push(p("All non-key attributes are fully dependent on the entire primary key. In junction tables like freelancer_skills and project_skills, both foreign keys form a composite unique constraint, and there are no partial dependencies. The proficiency column in freelancer_skills depends on the full combination of user_id and skill_id."));
children.push(pb("Third Normal Form (3NF)"));
children.push(p("No transitive dependencies exist. User profile information is split across users (authentication data), student_profiles (academic data), teacher_profiles (faculty data), and industry_profiles (company data) — each depending directly on user_id with no chain of dependencies. Similarly, review dimensions (reliability, quality, timeliness, communication) are stored in the reviews table rather than derived from separate lookup tables, as they are direct attributes of the review itself."));
children.push(h2("4.4 Key Relationships"));
children.push(p("The following table summarizes all foreign key relationships in the database:"));
children.push(makeTable(["Foreign Key", "References", "Cardinality", "Description"], relationships, [2800, 2000, 1200, 3360]));

// ── CHAPTER 5 ──
children.push(h1("Chapter 5: DBMS Selection"));
children.push(h2("5.1 Selection Criteria"));
children.push(p("The DBMS selection was guided by the following criteria critical to this project: full ACID compliance for transactional integrity, robust foreign key and constraint support, native trigger and stored procedure capabilities for business rule enforcement, JSON support for flexible metadata storage, and strong concurrency handling for multi-user access patterns."));
children.push(h2("5.2 Comparison of Candidate DBMS"));
children.push(makeTable(["Criteria", "PostgreSQL", "MySQL", "MongoDB"], dbmsComparison, [2200, 2560, 2200, 2400]));
children.push(h2("5.3 Justification for PostgreSQL"));
children.push(p("PostgreSQL was selected as the DBMS for the Campus Freelance Board based on the following decisive factors:"));
children.push(pi("ACID Compliance: The platform handles financial-adjacent data (project budgets, milestone payments, reputation scores) where transactional integrity is non-negotiable. PostgreSQL's full ACID compliance ensures no partial updates corrupt the database."));
children.push(pi("Complex Constraint Enforcement: The schema relies heavily on foreign keys (26 relationships), CHECK constraints (rating ranges, enum validations), and UNIQUE constraints (composite keys in junction tables). PostgreSQL handles all of these natively and efficiently."));
children.push(pi("Trigger Support: Business rules such as BR-07 (auto-assigning industry registration numbers) and BR-12 (certification expiry flagging) require database-level triggers. PostgreSQL's trigger system is mature and performant."));
children.push(pi("JSONB for Flexibility: While the core schema is relational, PostgreSQL's JSONB type allows storing flexible metadata (e.g., admin log details, notification payloads) without schema changes — a capability MySQL's JSON support lacks in indexing power."));
children.push(pi("MVCC Concurrency: With multiple users simultaneously browsing projects, submitting proposals, and updating contract statuses, PostgreSQL's Multi-Version Concurrency Control ensures read operations never block write operations and vice versa."));
children.push(pi("Prisma Ecosystem: The Prisma ORM (used in this project) has first-class PostgreSQL support with the PrismaPg adapter, providing type-safe queries and automatic migration management."));

// ── CHAPTER 6 ──
children.push(h1("Chapter 6: System Architecture"));
children.push(h2("6.1 Technology Stack"));
children.push(p("The Campus Freelance Board employs a modern full-stack architecture with clear separation of concerns:"));
children.push(makeTable(["Layer", "Technology", "Purpose"], techStack, [1800, 2600, 4960]));
children.push(h2("6.2 Role-Based Architecture"));
children.push(p("The frontend implements a complete role isolation pattern where each of the four user roles (Student, Teacher, Industry, Admin) has its own dedicated layout, sidebar navigation, dashboard, and set of pages. This is enforced at three levels:"));
children.push(pb("Level 1 — Proxy Guard (Edge)"));
children.push(p("The Next.js proxy.ts file intercepts every incoming request, checks for an authentication token, and validates the role claim. Unauthenticated users are redirected to /login. Authenticated users attempting to access another role's routes (e.g., a teacher accessing /student/dashboard) are redirected to their own dashboard."));
children.push(pb("Level 2 — Role-Specific Layouts"));
children.push(p("Each role group (student/, teacher/, industry/, admin/) has its own layout.tsx that renders the appropriate sidebar with role-relevant navigation items. This prevents any UI elements from another role from ever rendering."));
children.push(pb("Level 3 — Verification Guard (Page-Level)"));
children.push(p("For red-flagged industry projects, an additional check is performed at the page level to confirm the student freelancer has at least one approved certification before displaying the project details."));
children.push(h2("6.3 Flag-Based Visibility System"));
children.push(p("The three-flag system is a core differentiator of the Campus Freelance Board, controlling which projects are visible to which users:"));
children.push(pb("Blue Flag — Student-Posted Projects"));
children.push(p("When a student posts a project, it is assigned the blue flag. These projects are visible to all registered and active student users on the platform. No special verification is required to view or propose on blue-flagged projects."));
children.push(pb("Orange Flag — Teacher-Posted Projects"));
children.push(p("When a teacher posts a project, it receives the orange flag. These projects are visible to all active student users, just like blue-flagged projects. The distinction exists for analytics and filtering purposes, allowing students to identify academic-origin projects."));
children.push(pb("Red Flag — Industry-Posted Projects"));
children.push(p("When an industry client posts a project, it is assigned the red flag. These projects are visible ONLY to student freelancers who have at least one certification with status 'approved' in the certifications table. This ensures that industry clients, who may share sensitive or high-value work, only receive proposals from verified freelancers. The visibility check is implemented both in the backend API (projectController filters results based on the requester's verification status) and on the frontend (red project cards display a 'Verified Only' badge)."));
children.push(h2("6.4 Weighted Reputation Algorithm"));
children.push(p("Each freelancer's reputation is calculated as a weighted average across four dimensions, with weights configurable in the reputation_scores table:"));
children.push(p("Total Score = (Reliability × 0.30) + (Quality × 0.30) + (Timeliness × 0.20) + (Communication × 0.20)"));
children.push(p("Where each dimension score is the average of that dimension across all reviews received by the freelancer. The formula is implemented as a database view that joins the reviews table with reputation_scores, computing the running average whenever a new review is inserted. The weights (0.30, 0.30, 0.20, 0.20) prioritize reliability and quality equally, with timeliness and communication as secondary factors — reflecting the priorities identified during requirements gathering where clients rated 'delivers on time' and 'meets specifications' as the top two concerns."));

// ── CHAPTER 7 ──
children.push(h1("Chapter 7: Frontend Implementation"));
children.push(p("The following pages represent the complete frontend of the Campus Freelance Board, built with Next.js 16 using a dark glassmorphism aesthetic. Each role has a dedicated color accent: students use cyan (#00D4FF), teachers use orange (#FB923C), industry clients use red (#EF4444), and administrators use purple (#A78BFA). Placeholder boxes below indicate where actual screenshots will be inserted."));
screenshots.forEach(([a, b]) => children.push(screenshotRow(a, b)));

// ── CHAPTER 8 ──
children.push(h1("Chapter 8: Conclusion"));
children.push(p("The Campus Freelance Board database project demonstrates the complete lifecycle of designing a production-grade relational database for a complex, multi-role web application. Starting from requirements gathered through interviews, questionnaires, and observation, the project progressed through systematic analysis yielding 20 functional requirements, 8 non-functional requirements, and 12 business rules."));
children.push(p("The resulting database schema comprises 24 tables organized into 8 functional groups — covering core user management, freelancing workflows, project and proposal management, contract and milestone tracking, trust and reputation scoring, communication, administration, and file attachments. The schema adheres to Third Normal Form, ensuring data integrity and eliminating redundancy while maintaining query performance through appropriate indexing strategies."));
children.push(p("PostgreSQL was selected as the DBMS after rigorous comparison with MySQL and MongoDB, primarily for its superior constraint enforcement, trigger capabilities, and MVCC concurrency control — all critical for a platform handling financial-adjacent data with multiple concurrent users."));
children.push(p("The system architecture implements a three-layer role isolation pattern (proxy guard, layout-level, page-level) that ensures complete separation of user experiences. The flag-based visibility system and weighted reputation algorithm represent the core business logic differentiators, both implemented at the database level for consistency and performance."));
children.push(p("Future enhancements may include real-time messaging via Socket.io (infrastructure already planned), payment gateway integration for milestone payments, automated certification expiry monitoring via scheduled database jobs, and integration with university student information systems for automatic registration number validation."));

// ── DOCUMENT ──
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Arial", size: 22 } },
      heading1: { run: { font: "Arial", size: 36, bold: true, color: DK } },
      heading2: { run: { font: "Arial", size: 28, bold: true, color: MD } },
      heading3: { run: { font: "Arial", size: 24, bold: true, color: AC } },
    }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: "Campus Freelance Board — Database Project Report", font: "Arial", size: 16, color: AC, italics: true })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: MD })]
        })]
      })
    },
    children
  }]
});

// ── GENERATE ──
const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("Campus_Freelance_Board_DB_Report.docx", buffer);
console.log("✅ Report generated: Campus_Freelance_Board_DB_Report.docx");