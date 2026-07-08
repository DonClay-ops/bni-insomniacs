import { useState, useEffect, useRef, Fragment } from "react";
import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";

const supabase = createClient(
  process.env.NEXT_PUBLIC_SUPABASE_URL,
  process.env.NEXT_PUBLIC_SUPABASE_ANON_KEY
);

// ═══════════════════════════════════════════
// MEETING DATE — always the upcoming Wednesday (today, if today is Wednesday)
// ═══════════════════════════════════════════
const pad2 = (n) => String(n).padStart(2, "0");
const toYMD = (d) => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
const getNextWednesday = () => {
  const d = new Date();
  const diff = (3 - d.getDay() + 7) % 7; // 3 = Wednesday
  d.setDate(d.getDate() + diff);
  return toYMD(d);
};
const MEETING_DATE = getNextWednesday();

// ═══════════════════════════════════════════
// ASK LIFECYCLE — asks auto-archive after 6 weeks but stay available for AI matching
// ═══════════════════════════════════════════
const ASK_ACTIVE_DAYS = 42; // 6 weeks
const askAgeDays = (date) => Math.max(0, Math.floor((new Date(MEETING_DATE) - new Date(date)) / 86400000));
const isActiveAsk = (a) => a.status === "open" && askAgeDays(a.date) <= ASK_ACTIVE_DAYS;
const isArchivedAsk = (a) => a.status === "open" && askAgeDays(a.date) > ASK_ACTIVE_DAYS;

// ═══════════════════════════════════════════
// UAE OPEN CATEGORY REFERENCE — categories filled in other UAE BNI chapters but
// NOT held by any BNI Insomniacs member.
// Source: BNI_All_Members_Consolidated_v5.xlsx (37 UAE chapters, 1,085 members, Jul 2026).
// Format: [category, categoryGroup, chaptersFilledIn, membersNationwide], sorted by demand.
// The scan cross-checks this list against the LIVE member roster, so a category
// disappears automatically once an Insomniacs member fills it.
// ═══════════════════════════════════════════
const UAE_OPEN_CATEGORY_POOL = [
  ["CRM & ERP Solutions", "Computer & Programming", 17, 17],
  ["Executive Recruitment", "Employment Activities", 11, 11],
  ["Property Maintenance", "Real Estate Services", 10, 10],
  ["Fine Jewelry", "Retail & Wholesale", 9, 9],
  ["Gifts", "Retail & Wholesale", 8, 10],
  ["Financial Investments", "Finance & Insurance", 7, 7],
  ["Management Consulting", "Consulting", 7, 7],
  ["Sign Company", "Advertising & Marketing", 7, 7],
  ["IT Consultants", "Computer & Programming", 6, 8],
  ["Photography - Corporate Events, Wedding Etc", "Advertising & Marketing", 6, 6],
  ["Telecommunications Products/Services", "Telecommunications", 6, 6],
  ["Business Consultancy - Organization & Process", "Consulting", 5, 5],
  ["Interior Decorating", "Construction", 5, 5],
  ["Lighting Retailers", "Retail & Wholesale", 5, 5],
  ["Marketing Consultancy", "Advertising & Marketing", 5, 5],
  ["Medical Services", "Health & Wellness", 5, 5],
  ["Advertising Agency", "Advertising & Marketing", 4, 4],
  ["Event Planner", "Event & Business Service", 4, 4],
  ["HVAC - Heating & Air", "Construction", 4, 4],
  ["Litigation", "Legal & Accounting", 4, 4],
  ["App Developer", "Computer & Programming", 3, 3],
  ["Business Consultancy - Quality Management", "Consulting", 3, 3],
  ["Business Training/Coach", "Training & Coaching", 3, 3],
  ["Electrical Equipment", "Retail & Wholesale", 3, 3],
  ["Exhibitions, Conference & Seminar Organiser", "Event & Business Service", 3, 3],
  ["Furniture Retailer", "Retail & Wholesale", 3, 3],
  ["Government Services", "Legal & Accounting", 3, 3],
  ["Packaging", "Manufacturing", 3, 3],
  ["Printer - Large Format", "Advertising & Marketing", 3, 3],
  ["Stationery Supplies", "Retail & Wholesale", 3, 3],
  ["AI Consultant", "Advertising & Marketing", 2, 2],
  ["Artist", "Art & Entertainment", 2, 2],
  ["Auto/Car Rental/Leasing", "Automotives", 2, 2],
  ["Auto/Car Sales", "Automotives", 2, 2],
  ["Business Financing", "Finance & Insurance", 2, 2],
  ["CCTV", "Security & Investigation", 2, 2],
  ["Chemical Products", "Manufacturing", 2, 2],
  ["Chocolatier", "Food & Beverage", 2, 2],
  ["Clothing & Accessories", "Retail & Wholesale", 2, 2],
  ["Computer Security Solutions", "Computer & Programming", 2, 2],
  ["Computer Software", "Computer & Programming", 2, 2],
  ["Construction Project Management", "Construction", 2, 2],
  ["Custom Clothing/Tailor", "Retail & Wholesale", 2, 2],
  ["Diamonds & Gemstones", "Retail & Wholesale", 2, 2],
  ["E-Commerce Services", "Advertising & Marketing", 2, 2],
  ["Education Services/Tutor", "Training & Coaching", 2, 2],
  ["Educational Facility", "Training & Coaching", 2, 2],
  ["Elevators", "Construction", 2, 2],
  ["Event Rentals", "Event & Business Service", 2, 2],
  ["Fire Protection", "Security & Investigation", 2, 2],
  ["Flooring", "Construction", 2, 2],
  ["Geopathic Services", "Architecture & Engineering", 2, 2],
  ["Health & Wellness Products", "Health & Wellness", 2, 2],
  ["Joinery", "Construction", 2, 2],
  ["Landscape Maintenance & Supplies", "Architecture & Engineering", 2, 2],
  ["Lead Generation", "Advertising & Marketing", 2, 2],
  ["Life Coach", "Training & Coaching", 2, 2],
  ["Moving Company", "Transport & Shipping", 2, 2],
  ["Offshore Company Set up", "Legal & Accounting", 2, 2],
  ["Perfume", "Retail & Wholesale", 2, 2],
  ["Printing Products/Cartridges/Consumables", "Retail & Wholesale", 2, 2],
  ["Property Management", "Real Estate Services", 2, 2],
  ["Real Estate Development", "Real Estate Services", 2, 2],
  ["Security Products & Systems", "Security & Investigation", 2, 2],
  ["Steel Fabrication", "Manufacturing", 2, 2],
  ["Wills/Trusts", "Legal & Accounting", 2, 2],
  ["Apparel", "Manufacturing", 1, 1],
  ["Auto/Car Body Shop", "Automotives", 1, 1],
  ["Auto/Car Parts & Accessories", "Automotives", 1, 1],
  ["Automotive Expert", "Automotives", 1, 1],
  ["Banking Services", "Finance & Insurance", 1, 1],
  ["Business Consultancy - Small Business", "Consulting", 1, 1],
  ["Business Consultancy - Turnaround", "Consulting", 1, 1],
  ["Candles", "Retail & Wholesale", 1, 1],
  ["Chiropractor", "Health & Wellness", 1, 1],
  ["Citizenship Consultancy", "Legal & Accounting", 1, 1],
  ["Civil / Structural engineer", "Architecture & Engineering", 1, 1],
  ["Cleaning Products", "Retail & Wholesale", 1, 1],
  ["Cloud Services", "Computer & Programming", 1, 1],
  ["Commercial Builder", "Construction", 1, 1],
  ["Commercial Loans", "Finance & Insurance", 1, 1],
  ["Consumer Law", "Legal & Accounting", 1, 1],
  ["Copywriter & Writing Services", "Advertising & Marketing", 1, 1],
  ["Corporate Law", "Legal & Accounting", 1, 1],
  ["Cosmetics/Skin Care", "Personal Services", 1, 1],
  ["Counter Tops", "Construction", 1, 1],
  ["Credit Card/Merchant Services", "Finance & Insurance", 1, 1],
  ["Doctor/Physician", "Health & Wellness", 1, 1],
  ["Electronics Retailer", "Retail & Wholesale", 1, 1],
  ["Embroidery", "Advertising & Marketing", 1, 1],
  ["Environmental Services", "Construction", 1, 1],
  ["Estate Planning Law", "Legal & Accounting", 1, 1],
  ["Event Venue/Room Rental", "Event & Business Service", 1, 1],
  ["Events (Live)", "Event & Business Service", 1, 1],
  ["Flooring Retail", "Retail & Wholesale", 1, 1],
  ["Food Products", "Manufacturing", 1, 1],
  ["Foreign Exchange", "Finance & Insurance", 1, 1],
  ["Glass", "Construction", 1, 1],
  ["Health Facility/Gym/Club", "Health & Wellness", 1, 1],
  ["Home Automation", "Construction", 1, 1],
  ["Home Furnishings", "Retail & Wholesale", 1, 1],
  ["Home Staging", "Real Estate Services", 1, 1],
  ["IT Hardware Supplier", "Computer & Programming", 1, 1],
  ["In-Home Care", "Health & Wellness", 1, 1],
  ["Industrial Automation", "Architecture & Engineering", 1, 1],
  ["Insurance Consultant", "Finance & Insurance", 1, 1],
  ["Legal Service Plan", "Legal & Accounting", 1, 1],
  ["Machinery & Equipment Manufacture", "Manufacturing", 1, 1],
  ["Online Marketing", "Advertising & Marketing", 1, 1],
  ["Organizations & Other Specialist", "Organizations & Others", 1, 1],
  ["Painter - Residential", "Construction", 1, 1],
  ["Paper & Paper Products", "Manufacturing", 1, 1],
  ["Photography - Commercial/Industrial", "Advertising & Marketing", 1, 1],
  ["Print Advertising", "Advertising & Marketing", 1, 1],
  ["Printer - Digital", "Advertising & Marketing", 1, 1],
  ["Real Estate Inspector", "Real Estate Services", 1, 1],
  ["Relationship Marketing", "Advertising & Marketing", 1, 1],
  ["Salon/Spa", "Personal Services", 1, 1],
  ["Security Specialist", "Security & Investigation", 1, 1],
  ["Solar Systems", "Construction", 1, 1],
  ["Tiles", "Retail & Wholesale", 1, 1],
  ["Tires & Lubes", "Automotives", 1, 1],
  ["Translator/Language Services", "Event & Business Service", 1, 1],
  ["Water Systems", "Retail & Wholesale", 1, 1],
  ["Windows & Doors", "Construction", 1, 1]
];

// ═══════════════════════════════════════════
// SHARED CLAUDE API HELPER — correct browser headers, JSON response parsing
// ═══════════════════════════════════════════
const callClaude = async (prompt, maxTokens = 1500) => {
  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": process.env.NEXT_PUBLIC_ANTHROPIC_API_KEY,
      "anthropic-version": "2023-06-01",
      "anthropic-dangerous-direct-browser-access": "true",
    },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: maxTokens,
      messages: [{ role: "user", content: prompt }],
    }),
  });
  const data = await response.json();
  if (data.error) throw new Error(data.error.message || "API error");
  const text = (data.content || []).filter(b => b.type === "text").map(b => b.text).join("\n");
  return JSON.parse(text.replace(/```json|```/g, "").trim());
};

// Next N Wednesdays (starting with the upcoming one) — used for the template date dropdown
const upcomingWednesdays = (n = 8) => {
  const out = [];
  const d = new Date();
  d.setDate(d.getDate() + ((3 - d.getDay() + 7) % 7));
  for (let i = 0; i < n; i++) { out.push(toYMD(d)); d.setDate(d.getDate() + 7); }
  return out;
};

// Normalise dates coming from Excel: serial numbers, DD/MM/YYYY, YYYY-MM-DD, or blank
const normalizeExcelDate = (val) => {
  if (val === null || val === undefined || String(val).trim() === "") return MEETING_DATE;
  if (typeof val === "number") {
    // Excel serial date → JS date
    const d = new Date(Math.round((val - 25569) * 86400 * 1000));
    return isNaN(d) ? MEETING_DATE : toYMD(d);
  }
  const s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const dmy = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/); // DD/MM/YYYY (UAE format)
  if (dmy) {
    const yr = dmy[3].length === 2 ? `20${dmy[3]}` : dmy[3];
    return `${yr}-${pad2(+dmy[2])}-${pad2(+dmy[1])}`;
  }
  const d = new Date(s);
  return isNaN(d) ? MEETING_DATE : toYMD(d);
};

// ═══════════════════════════════════════════
// REAL MEMBER DATA - BNI INSOMNIACS  (now used as INITIAL state, editable)
// ═══════════════════════════════════════════
const INITIAL_MEMBERS = [
  { id:1, name:"A Syeeduddin", category:"Retail & Wholesale", specialty:"Retail/Wholesale Specialist" },
  { id:2, name:"Abhaysingh Chawan", category:"Event & Business Service", specialty:"Event & Business-Service Specialist" },
  { id:3, name:"Akshat Jain", category:"Manufacturing", specialty:"Furniture Manufacture" },
  { id:4, name:"Akshay Ramchandani", category:"Construction", specialty:"Renovations/Remodeling" },
  { id:5, name:"Akshay Revankar", category:"Transport & Shipping", specialty:"Transport & Shipping Specialist" },
  { id:6, name:"Amit Badlani", category:"Advertising & Marketing", specialty:"Printer - Offset" },
  { id:7, name:"Anand Bhaskar", category:"Consulting", specialty:"Management Consulting" },
  { id:8, name:"Ankit Bhansali", category:"Real Estate Services", specialty:"Real Estate Services Specialist" },
  { id:9, name:"Ankita Rao", category:"Architecture & Engineering", specialty:"Architect" },
  { id:10, name:"Anuj Malhotra", category:"Consulting", specialty:"Consulting Specialist" },
  { id:11, name:"Ashish Lalwani", category:"Real Estate Services", specialty:"Residential Real Estate - Sales" },
  { id:12, name:"Bharat Aidasani", category:"Travel", specialty:"Travel Agent - Outbound" },
  { id:13, name:"Bharat Karani", category:"Advertising & Marketing", specialty:"Advertising & Marketing Specialist" },
  { id:14, name:"Bhaskar Shah", category:"Computer & Programming", specialty:"Computer & Programming Specialist" },
  { id:15, name:"Bhawna Chopra", category:"Food & Beverage", specialty:"Food Services & Distributors" },
  { id:16, name:"Boman Parakh", category:"Construction", specialty:"Construction Specialist" },
  { id:17, name:"Chetan Shamji Thaleshwar", category:"Retail & Wholesale", specialty:"Fine Jewelry" },
  { id:18, name:"David Prabhu", category:"Automotives", specialty:"Automotive Specialist" },
  { id:19, name:"Deepak Bhagchandani", category:"Transport & Shipping", specialty:"Commercial Transportation" },
  { id:20, name:"Dharam Seth", category:"Employment Activities", specialty:"Employment Activities Specialist" },
  { id:21, name:"Emil Sunil George", category:"Legal & Accounting", specialty:"Legal Services - Commercial/Business" },
  { id:22, name:"Fariba Fattahi", category:"Health & Wellness", specialty:"Physiotherapist" },
  { id:23, name:"Fatma Siddiqui", category:"Event & Business Service", specialty:"Serviced Offices & Business Centre" },
  { id:24, name:"Garima Batra Juneja", category:"Real Estate Services", specialty:"Commercial Real Estate" },
  { id:25, name:"Girish Nathaney", category:"Advertising & Marketing", specialty:"Web Design & Development" },
  { id:26, name:"Govind Katara", category:"Advertising & Marketing", specialty:"Promotional Products" },
  { id:27, name:"Hakimuddin Saify", category:"Construction", specialty:"Construction Specialist" },
  { id:28, name:"Haresh Lalwani", category:"Finance & Insurance", specialty:"Residential Mortgages" },
  { id:29, name:"Hashem Mohamed Assaad", category:"Construction", specialty:"Builder/General Contractor" },
  { id:30, name:"Hemant Varandani", category:"Travel", specialty:"Travel Agent - Inbound" },
  { id:31, name:"Jacob Alex", category:"Food & Beverage", specialty:"Caterer" },
  { id:32, name:"Jai Satwani", category:"Construction", specialty:"Commercial/Retail Interior Design & Fitout" },
  { id:33, name:"Jasbir Bindra", category:"Legal & Accounting", specialty:"Legal & Accounting Specialist" },
  { id:34, name:"Jatin Sachdeva", category:"Health & Wellness", specialty:"Health & Wellness Specialist" },
  { id:35, name:"Jesika Menon", category:"Training & Coaching", specialty:"Learning Centre" },
  { id:36, name:"Karan Mulani", category:"Finance & Insurance", specialty:"General Insurance including Employee Benefits" },
  { id:37, name:"Leena Jayachandran", category:"Training & Coaching", specialty:"Sales Training/Coach" },
  { id:38, name:"Lijo Ittoop", category:"Advertising & Marketing", specialty:"Social Media" },
  { id:39, name:"Madhu Pallath", category:"Computer & Programming", specialty:"Computer & Programming Specialist" },
  { id:40, name:"Madhur Gupta", category:"Personal Services", specialty:"Personal Services Specialist" },
  { id:41, name:"Mahrukh Kazmi", category:"Legal & Accounting", specialty:"Intellectual Property Law" },
  { id:42, name:"Manoharan Chirakkal", category:"Computer & Programming", specialty:"Computer Networks & AMC" },
  { id:43, name:"Manoj Sureka", category:"Legal & Accounting", specialty:"Legal & Accounting Specialist" },
  { id:44, name:"Mariam Wehbi", category:"Legal & Accounting", specialty:"Taxation" },
  { id:45, name:"Mitun De Sarkar", category:"Health & Wellness", specialty:"Nutritionist" },
  { id:46, name:"Mohamed Faisal Ibrahimkutty", category:"Legal & Accounting", specialty:"Offshore Company Set up" },
  { id:47, name:"Mohit Sharma", category:"Finance & Insurance", specialty:"Wealth Management" },
  { id:48, name:"Mufaddal Boriyawala", category:"Retail & Wholesale", specialty:"Furniture Sales Commercial & Office" },
  { id:49, name:"Muhammad Salman", category:"Health & Wellness", specialty:"Personal Trainer - Fitness" },
  { id:50, name:"Mukeshkumar Ramani", category:"Health & Wellness", specialty:"Medical Services" },
  { id:51, name:"Muneer Samnani", category:"Training & Coaching", specialty:"Training & Coaching Specialist" },
  { id:52, name:"Nadeem Rasheed", category:"Retail & Wholesale", specialty:"Construction Products Retail" },
  { id:53, name:"Nitin Gupta", category:"Consulting", specialty:"Business Consulting" },
  { id:54, name:"Parichay Swarup", category:"Advertising & Marketing", specialty:"Brand Consultancy" },
  { id:55, name:"Paul Wesley", category:"Advertising & Marketing", specialty:"Television Advertising" },
  { id:56, name:"Peter Rodrigues", category:"Construction", specialty:"Construction Specialist" },
  { id:57, name:"Poonam Dabur", category:"Employment Activities", specialty:"Human Resources Consultancy" },
  { id:58, name:"Punit Thawani", category:"Health & Wellness", specialty:"General Dentist" },
  { id:59, name:"Rahul Datta", category:"Event & Business Service", specialty:"Corporate Events" },
  { id:60, name:"Rajesh Lobo", category:"Automotives", specialty:"Auto/Car Repair" },
  { id:61, name:"Rajesh Mirchandani", category:"Finance & Insurance", specialty:"Financial Planning & Personal Life Insurance" },
  { id:62, name:"Rajesh Pereira", category:"Event & Business Service", specialty:"Sound/Lighting/Staging" },
  { id:63, name:"Rakesh Pardasani", category:"Legal & Accounting", specialty:"Audit & Assurance" },
  { id:64, name:"Ranjeet Dang", category:"Retail & Wholesale", specialty:"Retail/Wholesale Specialist" },
  { id:65, name:"Rashida Malik Bathija", category:"Advertising & Marketing", specialty:"Advertising & Marketing Specialist" },
  { id:66, name:"Ritesh Rohra", category:"Retail & Wholesale", specialty:"Hardware Supplies" },
  { id:67, name:"Rutuja Marfatia", category:"Advertising & Marketing", specialty:"Public Relation" },
  { id:68, name:"Sachin Gupta", category:"Advertising & Marketing", specialty:"Search Engine Optimisation" },
  { id:69, name:"Sachin Singhal", category:"Personal Services", specialty:"Wedding Planner" },
  { id:70, name:"Sahil Gupta", category:"Advertising & Marketing", specialty:"Videographer/Film Producer" },
  { id:71, name:"Sanjay Nagdev", category:"Automotives", specialty:"Auto/Car Rental/Leasing" },
  { id:72, name:"Sanket Jain", category:"Transport & Shipping", specialty:"Shuttle/Limousine Service" },
  { id:73, name:"Satyanarayan Karan", category:"Retail & Wholesale", specialty:"Retail/Wholesale Specialist" },
  { id:74, name:"Saud Usman", category:"Retail & Wholesale", specialty:"Florist" },
  { id:75, name:"Saurabh Shetye", category:"Art & Entertainment", specialty:"Musicians" },
  { id:76, name:"Shahem Sabbagh", category:"Animals", specialty:"Veterinarian" },
  { id:77, name:"Shiva Purswani", category:"Retail & Wholesale", specialty:"Custom Apparel & Uniforms" },
  { id:78, name:"Shubhi Biju", category:"Architecture & Engineering", specialty:"Landscape Design & Contracting" },
  { id:79, name:"Simran Samtani", category:"Legal & Accounting", specialty:"Book Keeping & Out-Sourced CFO Services" },
  { id:80, name:"Sneha Bhatia", category:"Food & Beverage", specialty:"Cakes & Pastries" },
  { id:81, name:"Steve Cardoz", category:"Retail & Wholesale", specialty:"Curtains/Blinds" },
  { id:82, name:"Sumesh Wadhwa", category:"Manufacturing", specialty:"Manufacturing Specialist" },
  { id:83, name:"Sunil Gidhwani", category:"Real Estate Services", specialty:"Cleaning Service" },
  { id:84, name:"Sunil Padmanabhan", category:"Advertising & Marketing", specialty:"Advertising & Marketing Specialist" },
  { id:85, name:"Surjit Singh Namli", category:"Architecture & Engineering", specialty:"Residential Interior Design & Furniture" },
  { id:86, name:"Trupti Nilesh Rele", category:"Retail & Wholesale", specialty:"Art Dealer/Gallery Owner" },
  { id:87, name:"Umang Bhartia", category:"Transport & Shipping", specialty:"Courier Delivery Service" },
  { id:88, name:"Veena Muralidharan", category:"Transport & Shipping", specialty:"Freight Forwarding/Logistics" },
  { id:89, name:"Zankhana Mistry", category:"Health & Wellness", specialty:"Health & Wellness Services" },
];

// Master list of all known BNI categories so the user can pick from a sensible list
const ALL_BNI_CATEGORIES = [
  "Advertising & Marketing","Animals","Architecture & Engineering","Art & Entertainment","Automotives",
  "Computer & Programming","Construction","Consulting","Education","Employment Activities",
  "Event & Business Service","Finance & Insurance","Food & Beverage","Health & Wellness",
  "Legal & Accounting","Manufacturing","Personal Services","Real Estate Services","Retail & Wholesale",
  "Training & Coaching","Transport & Shipping","Travel"
].sort();

const INITIAL_ASKS = [
  { id: 1, memberId: 56, memberName: "Peter Rodrigues", askType: "specific_person", targetName: "Pallavi Dean", targetCompany: "Roar", targetCategory: "Architecture & Engineering", targetRole: "", notes: "Interior design project lead", date: "2026-03-10", status: "open" },
  { id: 2, memberId: 62, memberName: "Rajesh Pereira", askType: "specific_company", targetName: "", targetCompany: "Pixl Global", targetCategory: "Advertising & Marketing", targetRole: "", notes: "AV and staging partnership", date: "2026-03-10", status: "open" },
  { id: 3, memberId: 14, memberName: "Bhaskar Shah", askType: "general_role", targetName: "", targetCompany: "", targetCategory: "Computer & Programming", targetRole: "CFOs of companies looking for ERP solutions", notes: "Clay ERP implementation", date: "2026-03-10", status: "open" },
  { id: 4, memberId: 32, memberName: "Jai Satwani", askType: "general_role", targetName: "", targetCompany: "", targetCategory: "Construction", targetRole: "Hotel and restaurant owners planning renovations", notes: "Commercial fitout projects", date: "2026-03-03", status: "open" },
  { id: 5, memberId: 47, memberName: "Mohit Sharma", askType: "general_role", targetName: "", targetCompany: "", targetCategory: "Finance & Insurance", targetRole: "Business owners with AED 500K+ in savings", notes: "Wealth management consultation", date: "2026-03-03", status: "open" },
];

const STATUS_COLORS = {
  registered: { bg: "#EEF2FF", text: "#4338CA", label: "Registered" },
  called: { bg: "#FEF3C7", text: "#92400E", label: "Called" },
  confirmed: { bg: "#D1FAE5", text: "#065F46", label: "Confirmed" },
  attended: { bg: "#DBEAFE", text: "#1E40AF", label: "Attended" },
  oriented: { bg: "#EDE9FE", text: "#5B21B6", label: "Oriented" },
  applied: { bg: "#FCE7F3", text: "#9D174D", label: "Applied" },
  joined: { bg: "#D1FAE5", text: "#065F46", label: "Joined!" },
  declined: { bg: "#FEE2E2", text: "#991B1B", label: "Declined" },
  noshow: { bg: "#F3F4F6", text: "#6B7280", label: "No-Show" },
};

const INITIAL_VISITORS = [
  // ── Previous meetings ──
  { id: 4, name: "David Chen", business: "IT Solutions Provider", phone: "+971 56 321 0987", email: "david@example.com", invitedBy: "Bhaskar Shah", status: "attended", callNotes: "Has 50+ business cards", category: "Computer & Programming", specialty: "IT Solutions", seatAssignment: "Next to Manoharan (Networks)", followUpResponse: "questions", date: "2026-03-10", bio: null, validation: null },
  { id: 5, name: "Amina Yusuf", business: "Legal Consultancy", phone: "+971 54 654 3210", email: "amina@example.com", invitedBy: "Emil Sunil George", status: "applied", callNotes: "Returning visitor, very engaged", category: "Legal & Accounting", specialty: "Corporate Law", seatAssignment: "Next to Mahrukh (IP Law)", followUpResponse: "ready", date: "2026-03-10", bio: null, validation: null },
  // ── Demo visitors (replaced by Supabase data on load) ──
  { id: 10, name: "Ram Bahin (Substitute)", business: "Paparazzi House", phone: "050 8500750", email: "Director@paparazzi.house", invitedBy: "Anand Bhaskar", status: "confirmed", callNotes: "Substitute visitor", category: "Training & Coaching", specialty: "Business Training / Coach", seatAssignment: "", followUpResponse: null, date: MEETING_DATE, bio: null, validation: null },
  { id: 11, name: "Shahnawaz", business: "Dimos Café & Restaurant", phone: "055 8691155", email: "dimoscafeandrestaurant@gmail.com", invitedBy: "Ankita Rao", status: "called", callNotes: "Was busy. Details sent. Visit to be confirmed.", category: "Food & Beverage", specialty: "Restaurant / Café", seatAssignment: "", followUpResponse: null, date: MEETING_DATE, bio: null, validation: null },
  { id: 12, name: "Kapardhi Dhavala", business: "Property Maintainace", phone: "056 6817669", email: "kapardhi.dhavala@spacemanager.ae", invitedBy: "Ankita Rao", status: "registered", callNotes: "No. not reachable — Ankita informed.", category: "Construction", specialty: "Property Maintenance", seatAssignment: "", followUpResponse: null, date: MEETING_DATE, bio: null, validation: null },
  { id: 13, name: "Kevin Monteiro", business: "Voxtel Communication", phone: "055 1226166", email: "kevin@voxtelme.com", invitedBy: "", status: "registered", callNotes: "No. not reachable or constantly busy.", category: "Computer & Programming", specialty: "IT Consultants / Communication", seatAssignment: "", followUpResponse: null, date: MEETING_DATE, bio: null, validation: null },
  { id: 14, name: "Varaprasad SN", business: "Optculture", phone: "052 9045457", email: "varaprasad@optculture.com", invitedBy: "Madhu Pallath", status: "confirmed", callNotes: "Confirmed. Details sent.", category: "Advertising & Marketing", specialty: "Customer Loyalty & Engagement", seatAssignment: "", followUpResponse: null, date: MEETING_DATE, bio: null, validation: null },
  { id: 15, name: "Kanchan", business: "KLIPIT", phone: "052 254 6953", email: "kanchan.magaji@klipit.co", invitedBy: "Madhu Pallath", status: "registered", callNotes: "No. not reachable. Madhu informed.", category: "Advertising & Marketing", specialty: "Digital Reimbursement Platform", seatAssignment: "", followUpResponse: null, date: MEETING_DATE, bio: null, validation: null },
];

// ═══════════════════════════════════════════
// AI MATCHING ENGINE
// ═══════════════════════════════════════════
function findMatches(visitor, asks, members) {
  const matches = [];
  const vCat = (visitor.category || "").toLowerCase();
  const vSpec = (visitor.specialty || "").toLowerCase();
  const vBiz = (visitor.business || "").toLowerCase();
  const vName = (visitor.name || "").toLowerCase();

  asks.filter(a => a.status === "open").forEach(ask => {
    let score = 0;
    let reason = "";
    const aCat = (ask.targetCategory || "").toLowerCase();
    const aRole = (ask.targetRole || "").toLowerCase();
    const aCompany = (ask.targetCompany || "").toLowerCase();
    const aName = (ask.targetName || "").toLowerCase();

    if (ask.askType === "specific_person" && aName && vName.includes(aName.split(" ")[0].toLowerCase())) {
      score = 100; reason = `Exact person match: ${ask.memberName} is looking for ${ask.targetName}`;
    } else if (ask.askType === "specific_company" && aCompany && vBiz.toLowerCase().includes(aCompany.toLowerCase())) {
      score = 95; reason = `Company match: ${ask.memberName} is looking for someone from ${ask.targetCompany}`;
    } else if (aCat && vCat === aCat) {
      score = 70; reason = `Category match: ${ask.memberName} is looking for ${ask.targetRole || ask.targetCategory}`;
      if (aRole && (vSpec.includes(aRole.split(" ")[0].toLowerCase()) || vBiz.includes(aRole.split(" ")[0].toLowerCase()))) {
        score = 85; reason = `Strong match: ${ask.memberName} is looking for "${ask.targetRole}" — visitor's business aligns`;
      }
    }
    if (score > 0) matches.push({ type: "ask", score, reason, member: members.find(m => m.id === ask.memberId), ask });
  });

  members.filter(m => m.category.toLowerCase() === vCat).forEach(m => {
    if (!matches.find(mt => mt.member?.id === m.id)) {
      matches.push({ type: "contact_sphere", score: 50, reason: `Contact Sphere: ${m.name} (${m.specialty}) is in the same category`, member: m });
    }
  });

  return matches.sort((a, b) => b.score - a.score);
}

// ═══════════════════════════════════════════
// UI COMPONENTS
// ═══════════════════════════════════════════
function Badge({ bg, text, label }) {
  return <span style={{ background: bg, color: text, padding: "3px 10px", borderRadius: 20, fontSize: 11, fontWeight: 600, whiteSpace: "nowrap" }}>{label}</span>;
}
function StatusBadge({ status }) {
  const s = STATUS_COLORS[status] || STATUS_COLORS.registered;
  return <Badge bg={s.bg} text={s.text} label={s.label} />;
}
function Card({ children, style, onClick }) {
  return <div onClick={onClick} style={{ background: "#fff", borderRadius: 12, border: "1px solid #E5E7EB", padding: 16, ...style, cursor: onClick ? "pointer" : "default" }}>{children}</div>;
}
function StatCard({ label, value, sub, color }) {
  return <Card style={{ flex: 1, minWidth: 130 }}>
    <div style={{ fontSize: 10, color: "#6B7280", textTransform: "uppercase", letterSpacing: 1, marginBottom: 2 }}>{label}</div>
    <div style={{ fontSize: 28, fontWeight: 800, color: color || "#111", lineHeight: 1.1 }}>{value}</div>
    {sub && <div style={{ fontSize: 11, color: "#9CA3AF", marginTop: 3 }}>{sub}</div>}
  </Card>;
}

// Confirmation modal — reused for delete actions
function ConfirmModal({ open, title, message, confirmLabel = "Delete", cancelLabel = "Cancel", danger = true, onConfirm, onCancel }) {
  if (!open) return null;
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(15,23,42,0.55)", zIndex: 10000, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
      <div style={{ background: "#fff", borderRadius: 14, padding: 22, maxWidth: 420, width: "100%", boxShadow: "0 20px 50px rgba(0,0,0,0.3)" }}>
        <div style={{ fontSize: 16, fontWeight: 800, color: danger ? "#991B1B" : "#1B2A4A", marginBottom: 8 }}>{title}</div>
        <div style={{ fontSize: 13, color: "#374151", lineHeight: 1.6, marginBottom: 16 }}>{message}</div>
        <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
          <button onClick={onCancel} style={{ padding: "8px 16px", borderRadius: 8, border: "1px solid #D1D5DB", background: "#fff", fontSize: 12, fontWeight: 600, cursor: "pointer", color: "#374151" }}>{cancelLabel}</button>
          <button onClick={onConfirm} style={{ padding: "8px 16px", borderRadius: 8, border: "none", background: danger ? "#8B1A1A" : "#1B2A4A", color: "#fff", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>{confirmLabel}</button>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════
// AI VISITOR INTELLIGENCE COMPONENT  (briefing card — unchanged from v1)
// ═══════════════════════════════════════════
function VisitorIntelligence({ visitor, onBioSaved }) {
  const [loading, setLoading] = useState(false);
  const [bio, setBio] = useState(visitor.bio || null);
  const [error, setError] = useState(null);

  const generateBio = async () => {
    setLoading(true);
    setError(null);
    try {
      const prompt = `You are a BNI chapter intelligence assistant helping the Visitor Host prepare warm, specific introductions for BNI Insomniacs Dubai.

Research this visitor and write a concise VIP briefing card. Use your knowledge of this business/industry to make it specific and useful — not generic.

Visitor Name: ${visitor.name}
Business Name: ${visitor.business}
Industry: ${visitor.category} — ${visitor.specialty}
Email: ${visitor.email || "not provided"}
Invited by BNI member: ${visitor.invitedBy || "walk-in"}
Call notes: ${visitor.callNotes || "None"}

Write a briefing in this exact JSON structure (no markdown, pure JSON):
{
  "headline": "One sharp sentence describing who they are and what they do — make it vivid and specific to their actual business, not generic",
  "businessSnapshot": "2-3 sentences about this specific business (${visitor.business}), their likely services, target clients, and how they operate in Dubai/UAE market",
  "conversationStarters": ["specific opener 1 relevant to their actual business", "specific opener 2 that references their industry context in UAE", "specific opener 3 about their growth or challenges"],
  "whatTheyNeed": "What business challenges or referrals someone running ${visitor.business} in the ${visitor.specialty} space likely needs — be specific",
  "whyBNI": "One compelling sentence on why BNI specifically benefits ${visitor.business} — mention referral potential from relevant BNI categories",
  "introScript": "A warm 2-sentence introduction the Visitor Host could say when introducing ${visitor.name} from ${visitor.business} to a BNI member. Make it specific to their business.",
  "linkedinTip": "Exact search string to use on LinkedIn or Google to find ${visitor.name} at ${visitor.business} and what to look for"
}`;

      const parsed = await callClaude(prompt, 1000);
      setBio(parsed);
      onBioSaved(visitor.id, parsed);
    } catch (e) {
      setError("Could not generate bio. Please try again.");
    }
    setLoading(false);
  };

  if (!bio && !loading) return (
    <button onClick={generateBio} style={{
      background: "linear-gradient(135deg, #1B2A4A, #8B1A1A)", color: "#fff",
      border: "none", borderRadius: 8, padding: "8px 16px", fontSize: 12,
      fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", gap: 6
    }}>
      ✨ Generate VIP Briefing
    </button>
  );

  if (loading) return (
    <div style={{ display: "flex", alignItems: "center", gap: 8, color: "#6B7280", fontSize: 12, padding: "8px 0" }}>
      <div style={{ width: 16, height: 16, border: "2px solid #8B1A1A", borderTopColor: "transparent", borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
      Researching {visitor.name}...
    </div>
  );

  if (error) return <div style={{ color: "#991B1B", fontSize: 12 }}>{error} <button onClick={generateBio} style={{ textDecoration: "underline", background: "none", border: "none", cursor: "pointer", color: "#991B1B" }}>Retry</button></div>;

  return (
    <div style={{ marginTop: 10 }}>
      <div style={{ background: "linear-gradient(135deg, #0F172A 0%, #1B2A4A 50%, #2D1515 100%)", borderRadius: 10, padding: 14, color: "#fff" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
          <div>
            <div style={{ fontSize: 13, fontWeight: 800, color: "#fff", letterSpacing: -0.3 }}>✦ VIP Intelligence Briefing</div>
            <div style={{ fontSize: 11, color: "#94A3B8", marginTop: 2 }}>{bio.headline}</div>
          </div>
          <button onClick={() => { setBio(null); }} style={{ background: "rgba(255,255,255,0.1)", border: "none", borderRadius: 6, padding: "3px 8px", fontSize: 10, color: "#94A3B8", cursor: "pointer" }}>↺ Refresh</button>
        </div>
        <div style={{ background: "rgba(255,255,255,0.07)", borderRadius: 8, padding: 10, marginBottom: 8 }}>
          <div style={{ fontSize: 9, textTransform: "uppercase", letterSpacing: 1, color: "#64748B", marginBottom: 4 }}>Business Snapshot</div>
          <div style={{ fontSize: 11, color: "#CBD5E1", lineHeight: 1.6 }}>{bio.businessSnapshot}</div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
          <div style={{ background: "rgba(255,255,255,0.07)", borderRadius: 8, padding: 10 }}>
            <div style={{ fontSize: 9, textTransform: "uppercase", letterSpacing: 1, color: "#64748B", marginBottom: 4 }}>What They Need</div>
            <div style={{ fontSize: 11, color: "#CBD5E1", lineHeight: 1.5 }}>{bio.whatTheyNeed}</div>
          </div>
          <div style={{ background: "rgba(255,255,255,0.07)", borderRadius: 8, padding: 10 }}>
            <div style={{ fontSize: 9, textTransform: "uppercase", letterSpacing: 1, color: "#64748B", marginBottom: 4 }}>Why BNI for Them</div>
            <div style={{ fontSize: 11, color: "#CBD5E1", lineHeight: 1.5 }}>{bio.whyBNI}</div>
          </div>
        </div>
        <div style={{ background: "rgba(255,255,255,0.07)", borderRadius: 8, padding: 10, marginBottom: 8 }}>
          <div style={{ fontSize: 9, textTransform: "uppercase", letterSpacing: 1, color: "#64748B", marginBottom: 6 }}>💬 Conversation Starters</div>
          <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
            {bio.conversationStarters?.map((cs, i) => (
              <div key={i} style={{ display: "flex", gap: 6, alignItems: "flex-start" }}>
                <span style={{ color: "#8B1A1A", fontSize: 10, fontWeight: 800, flexShrink: 0, marginTop: 1 }}>→</span>
                <span style={{ fontSize: 11, color: "#CBD5E1", lineHeight: 1.5 }}>{cs}</span>
              </div>
            ))}
          </div>
        </div>
        <div style={{ background: "rgba(139,26,26,0.25)", border: "1px solid rgba(139,26,26,0.4)", borderRadius: 8, padding: 10, marginBottom: 8 }}>
          <div style={{ fontSize: 9, textTransform: "uppercase", letterSpacing: 1, color: "#F87171", marginBottom: 4 }}>🎤 Introduction Script</div>
          <div style={{ fontSize: 12, color: "#fff", lineHeight: 1.6, fontStyle: "italic" }}>"{bio.introScript}"</div>
        </div>
        <div style={{ display: "flex", gap: 6, alignItems: "flex-start" }}>
          <span style={{ fontSize: 10, color: "#60A5FA", flexShrink: 0 }}>🔍</span>
          <div style={{ fontSize: 10, color: "#60A5FA", lineHeight: 1.5 }}><strong>Research tip:</strong> {bio.linkedinTip}</div>
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════
// NEW: AI VISITOR VALIDATION COMPONENT
// Checks suitability — category conflicts with current members, red flags, fit score
// ═══════════════════════════════════════════
function VisitorValidation({ visitor, members, onValidationSaved }) {
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(visitor.validation || null);
  const [error, setError] = useState(null);

  const validate = async () => {
    setLoading(true);
    setError(null);
    try {
      // Build a snapshot of current members in the visitor's category — this is the conflict check
      const sameCategoryMembers = members
        .filter(m => m.category.toLowerCase() === (visitor.category || "").toLowerCase())
        .map(m => `${m.name} — ${m.specialty}`);

      const allCategoriesInChapter = [...new Set(members.map(m => m.category))].sort();

      const prompt = `You are the BNI Insomniacs Dubai Membership Committee assistant. Your job is to validate whether a prospective visitor is a SUITABLE FIT for the chapter, based on BNI's "one person per professional classification" rule and general chapter health.

VISITOR DETAILS:
- Name: ${visitor.name}
- Business: ${visitor.business}
- Category: ${visitor.category || "NOT SPECIFIED"}
- Specialty / Classification: ${visitor.specialty || "NOT SPECIFIED"}
- Invited by: ${visitor.invitedBy || "walk-in (not invited)"}
- Phone: ${visitor.phone || "not provided"}
- Email: ${visitor.email || "not provided"}
- Notes from intake call: ${visitor.callNotes || "none"}

CURRENT CHAPTER MEMBERS IN VISITOR'S CATEGORY (potential conflicts):
${sameCategoryMembers.length > 0 ? sameCategoryMembers.join("\n") : "(none — category is OPEN in the chapter)"}

ALL CATEGORIES PRESENT IN CHAPTER:
${allCategoriesInChapter.join(", ")}

Evaluate the visitor and return ONLY valid JSON in this exact structure (no markdown, no code fences):
{
  "verdict": "GREEN" | "AMBER" | "RED",
  "fitScore": <number 0-100>,
  "headline": "One short sentence summarising the verdict",
  "classificationCheck": {
    "status": "OPEN" | "CONFLICT" | "OVERLAP_RISK",
    "explanation": "Specific explanation. If OPEN, say no member currently holds this classification. If CONFLICT, name the member(s) who already hold it. If OVERLAP_RISK, explain the partial overlap and how it might be resolved (e.g., narrowing the specialty)."
  },
  "strengths": ["specific positive 1", "specific positive 2", "specific positive 3"],
  "concerns": ["specific concern 1 if any", "specific concern 2 if any"],
  "redFlags": ["actual red flags only — leave empty array if none"],
  "referralPotential": "Realistic 1-2 sentence assessment of how much referral business this visitor's category could generate FROM and TO the existing chapter members",
  "recommendation": "One of: 'Strongly recommend inviting back', 'Recommend inviting back', 'Invite back with caution — discuss classification first', 'Do not invite back', 'Needs more information'",
  "nextSteps": ["specific action 1 for the Visitor Host or Membership Committee", "specific action 2"]
}

Be honest and specific. If information is missing (no category, no business name, no inviter), flag it. If there is a clear classification conflict with an existing member, the verdict must be RED or AMBER — never GREEN.`;

      const parsed = await callClaude(prompt, 1200);
      setResult(parsed);
      onValidationSaved(visitor.id, parsed);
    } catch (e) {
      setError("Could not validate. Please try again.");
    }
    setLoading(false);
  };

  const verdictStyle = (v) => {
    if (v === "GREEN") return { bg: "#D1FAE5", border: "#10B981", text: "#065F46", icon: "✅" };
    if (v === "AMBER") return { bg: "#FEF3C7", border: "#F59E0B", text: "#92400E", icon: "⚠️" };
    if (v === "RED") return { bg: "#FEE2E2", border: "#EF4444", text: "#991B1B", icon: "🛑" };
    return { bg: "#F3F4F6", border: "#9CA3AF", text: "#374151", icon: "❓" };
  };

  if (!result && !loading) return (
    <button onClick={validate} style={{
      background: "linear-gradient(135deg, #047857, #065F46)", color: "#fff",
      border: "none", borderRadius: 8, padding: "8px 16px", fontSize: 12,
      fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", gap: 6
    }}>
      🛡️ Validate Visitor Fit
    </button>
  );

  if (loading) return (
    <div style={{ display: "flex", alignItems: "center", gap: 8, color: "#065F46", fontSize: 12, padding: "8px 0" }}>
      <div style={{ width: 16, height: 16, border: "2px solid #047857", borderTopColor: "transparent", borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
      Validating fit against chapter…
    </div>
  );

  if (error) return <div style={{ color: "#991B1B", fontSize: 12 }}>{error} <button onClick={validate} style={{ textDecoration: "underline", background: "none", border: "none", cursor: "pointer", color: "#991B1B" }}>Retry</button></div>;

  const vs = verdictStyle(result.verdict);

  return (
    <div style={{ marginTop: 10 }}>
      <div style={{ background: vs.bg, border: `2px solid ${vs.border}`, borderRadius: 12, padding: 14 }}>
        {/* Header */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10, marginBottom: 10 }}>
          <div style={{ flex: 1 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 3 }}>
              <span style={{ fontSize: 18 }}>{vs.icon}</span>
              <span style={{ fontSize: 14, fontWeight: 900, color: vs.text, letterSpacing: -0.3 }}>
                Verdict: {result.verdict}
              </span>
              <span style={{ background: "#fff", color: vs.text, fontSize: 10, fontWeight: 800, padding: "2px 8px", borderRadius: 12, border: `1px solid ${vs.border}` }}>
                Fit Score: {result.fitScore}/100
              </span>
            </div>
            <div style={{ fontSize: 12, color: vs.text, lineHeight: 1.5, fontWeight: 600 }}>{result.headline}</div>
          </div>
          <button onClick={() => { setResult(null); onValidationSaved(visitor.id, null); }} style={{ background: "#fff", border: `1px solid ${vs.border}`, borderRadius: 6, padding: "3px 8px", fontSize: 10, color: vs.text, cursor: "pointer", flexShrink: 0 }}>↺ Re-run</button>
        </div>

        {/* Classification check — most important block */}
        <div style={{ background: "#fff", borderRadius: 8, padding: 10, marginBottom: 8, borderLeft: `4px solid ${vs.border}` }}>
          <div style={{ fontSize: 9, textTransform: "uppercase", letterSpacing: 1, color: "#6B7280", marginBottom: 4, fontWeight: 700 }}>Classification Check</div>
          <div style={{ fontSize: 12, fontWeight: 800, color: result.classificationCheck?.status === "OPEN" ? "#065F46" : result.classificationCheck?.status === "CONFLICT" ? "#991B1B" : "#92400E", marginBottom: 3 }}>
            {result.classificationCheck?.status === "OPEN" && "✓ Category is OPEN"}
            {result.classificationCheck?.status === "CONFLICT" && "✗ CONFLICT with existing member"}
            {result.classificationCheck?.status === "OVERLAP_RISK" && "⚠ Overlap risk"}
          </div>
          <div style={{ fontSize: 11, color: "#374151", lineHeight: 1.6 }}>{result.classificationCheck?.explanation}</div>
        </div>

        {/* Two-column: Strengths + Concerns */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
          <div style={{ background: "#fff", borderRadius: 8, padding: 10 }}>
            <div style={{ fontSize: 9, textTransform: "uppercase", letterSpacing: 1, color: "#065F46", marginBottom: 4, fontWeight: 700 }}>👍 Strengths</div>
            {result.strengths?.length > 0 ? result.strengths.map((s, i) => (
              <div key={i} style={{ fontSize: 11, color: "#374151", lineHeight: 1.5, marginBottom: 3, display: "flex", gap: 5 }}>
                <span style={{ color: "#10B981", fontWeight: 800 }}>+</span><span>{s}</span>
              </div>
            )) : <div style={{ fontSize: 11, color: "#9CA3AF" }}>—</div>}
          </div>
          <div style={{ background: "#fff", borderRadius: 8, padding: 10 }}>
            <div style={{ fontSize: 9, textTransform: "uppercase", letterSpacing: 1, color: "#92400E", marginBottom: 4, fontWeight: 700 }}>⚠ Concerns</div>
            {result.concerns?.length > 0 ? result.concerns.map((c, i) => (
              <div key={i} style={{ fontSize: 11, color: "#374151", lineHeight: 1.5, marginBottom: 3, display: "flex", gap: 5 }}>
                <span style={{ color: "#F59E0B", fontWeight: 800 }}>!</span><span>{c}</span>
              </div>
            )) : <div style={{ fontSize: 11, color: "#9CA3AF" }}>None flagged</div>}
          </div>
        </div>

        {/* Red flags — only if any */}
        {result.redFlags?.length > 0 && (
          <div style={{ background: "#fff", border: "1px solid #EF4444", borderRadius: 8, padding: 10, marginBottom: 8 }}>
            <div style={{ fontSize: 9, textTransform: "uppercase", letterSpacing: 1, color: "#991B1B", marginBottom: 4, fontWeight: 700 }}>🚩 Red Flags</div>
            {result.redFlags.map((rf, i) => (
              <div key={i} style={{ fontSize: 11, color: "#991B1B", lineHeight: 1.5, marginBottom: 2 }}>• {rf}</div>
            ))}
          </div>
        )}

        {/* Referral potential */}
        <div style={{ background: "#fff", borderRadius: 8, padding: 10, marginBottom: 8 }}>
          <div style={{ fontSize: 9, textTransform: "uppercase", letterSpacing: 1, color: "#6B7280", marginBottom: 4, fontWeight: 700 }}>💼 Referral Potential</div>
          <div style={{ fontSize: 11, color: "#374151", lineHeight: 1.6 }}>{result.referralPotential}</div>
        </div>

        {/* Recommendation — highlighted */}
        <div style={{ background: vs.text, color: "#fff", borderRadius: 8, padding: 10, marginBottom: 8 }}>
          <div style={{ fontSize: 9, textTransform: "uppercase", letterSpacing: 1, color: "rgba(255,255,255,0.7)", marginBottom: 3, fontWeight: 700 }}>📋 Recommendation</div>
          <div style={{ fontSize: 13, fontWeight: 800, lineHeight: 1.5 }}>{result.recommendation}</div>
        </div>

        {/* Next steps */}
        <div style={{ background: "#fff", borderRadius: 8, padding: 10 }}>
          <div style={{ fontSize: 9, textTransform: "uppercase", letterSpacing: 1, color: "#6B7280", marginBottom: 4, fontWeight: 700 }}>👉 Next Steps</div>
          {result.nextSteps?.map((ns, i) => (
            <div key={i} style={{ fontSize: 11, color: "#374151", lineHeight: 1.5, marginBottom: 3, display: "flex", gap: 5 }}>
              <span style={{ color: vs.border, fontWeight: 800 }}>{i + 1}.</span><span>{ns}</span>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════
// SEAT PLANNER COMPONENT  (unchanged from v1, takes members as prop now)
// ═══════════════════════════════════════════
function SeatPlanner({ visitors, asks, members }) {
  const meetingVisitors = visitors.filter(v => v.date === MEETING_DATE);
  const [seats, setSeats] = useState(() => {
    const assigned = {};
    meetingVisitors.slice(0, 6).forEach((v, i) => { assigned[`V${i + 1}`] = v.id; });
    return assigned;
  });
  const [dragging, setDragging] = useState(null);
  const [hovering, setHovering] = useState(null);

  const autoAssign = () => {
    const memberSeats = {
      M1: members.find(m => m.category === "Finance & Insurance"),
      M2: members.find(m => m.category === "Legal & Accounting"),
      M3: members.find(m => m.category === "Advertising & Marketing"),
      M4: members.find(m => m.category === "Construction"),
      M5: members.find(m => m.category === "Health & Wellness"),
      M6: members.find(m => m.category === "Real Estate Services"),
    };
    const newSeats = {};
    const assigned = new Set();
    ["V1","V2","V3","V4","V5","V6"].forEach((seatKey, i) => {
      const adjacentMember = Object.values(memberSeats)[i];
      if (!adjacentMember) return;
      let bestVisitor = null;
      let bestScore = -1;
      meetingVisitors.forEach(v => {
        if (assigned.has(v.id)) return;
        const matches = findMatches(v, asks, members);
        const memberMatch = matches.find(m => m.member?.id === adjacentMember.id);
        const catMatch = v.category === adjacentMember.category ? 60 : 0;
        const score = (memberMatch?.score || 0) + catMatch;
        if (score > bestScore) { bestScore = score; bestVisitor = v; }
      });
      if (!bestVisitor && meetingVisitors.length > 0) {
        bestVisitor = meetingVisitors.find(v => !assigned.has(v.id)) || null;
      }
      if (bestVisitor) {
        newSeats[seatKey] = bestVisitor.id;
        assigned.add(bestVisitor.id);
      }
    });
    setSeats(newSeats);
  };

  const handleDrop = (targetSeat) => {
    if (!dragging) return;
    setSeats(prev => {
      const next = { ...prev };
      const fromSeat = Object.entries(next).find(([, vid]) => vid === dragging)?.[0];
      const displaced = next[targetSeat];
      if (fromSeat) next[fromSeat] = displaced || null;
      next[targetSeat] = dragging;
      return next;
    });
    setDragging(null);
    setHovering(null);
  };

  const getVisitor = (seatKey) => meetingVisitors.find(v => v.id === seats[seatKey]);
  const unassigned = meetingVisitors.filter(v => !Object.values(seats).includes(v.id));

  const catColor = (cat) => {
    const map = {
      "Finance & Insurance": "#1D4ED8","Legal & Accounting": "#7C3AED","Advertising & Marketing": "#D97706",
      "Construction": "#B45309","Health & Wellness": "#059669","Real Estate Services": "#0891B2",
      "Computer & Programming": "#6D28D9","Transport & Shipping": "#374151","Consulting": "#7C3AED",
      "Event & Business Service": "#BE185D","Training & Coaching": "#15803D","Travel": "#0369A1",
      "Food & Beverage": "#B91C1C","Retail & Wholesale": "#92400E","Architecture & Engineering": "#0F766E",
      "Manufacturing": "#6B21A8","Personal Services": "#BE185D","Automotives": "#374151","Employment Activities": "#1E40AF",
    };
    return map[cat] || "#6B7280";
  };

  const visitorSeats = ["V1","V2","V3","V4","V5","V6"];
  const displayMembers = [
    members.find(m => m.specialty === "Wealth Management"),
    members.find(m => m.specialty === "Residential Mortgages"),
    members.find(m => m.specialty === "General Insurance including Employee Benefits"),
    members.find(m => m.specialty === "Commercial Real Estate"),
    members.find(m => m.specialty === "Architect"),
    members.find(m => m.specialty === "Social Media"),
    members.find(m => m.specialty === "Search Engine Optimisation"),
    members.find(m => m.specialty === "Commercial/Retail Interior Design & Fitout"),
    members.find(m => m.specialty === "Builder/General Contractor"),
    members.find(m => m.specialty === "Construction Specialist"),
  ].filter(Boolean);

  const SeatChip = ({ seatKey, isVisitor }) => {
    const v = isVisitor ? getVisitor(seatKey) : null;
    const m = !isVisitor ? displayMembers[parseInt(seatKey.replace("M","")) - 1] : null;
    const isOver = hovering === seatKey;
    return (
      <div onDragOver={(e) => { e.preventDefault(); setHovering(seatKey); }} onDragLeave={() => setHovering(null)} onDrop={() => isVisitor && handleDrop(seatKey)}
        style={{ width: isVisitor ? 90 : 82, minHeight: isVisitor ? 72 : 64, borderRadius: 8, border: isOver ? "2px dashed #8B1A1A" : isVisitor ? "2px solid #8B1A1A" : "2px solid #1B2A4A", background: isOver ? "#FFF1F1" : isVisitor ? (v ? "#FFF7F7" : "#FFF1F1") : (m ? "#EEF2FF" : "#F8FAFC"), padding: "6px 5px", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", transition: "all 0.15s", position: "relative" }}>
        <div style={{ position: "absolute", top: 3, left: 5, fontSize: 8, fontWeight: 800, color: isVisitor ? "#8B1A1A" : "#1B2A4A", opacity: 0.6 }}>{seatKey}</div>
        {isVisitor && v && (
          <>
            <div draggable onDragStart={() => setDragging(v.id)} style={{ cursor: "grab" }}>
              <div style={{ width: 28, height: 28, borderRadius: "50%", background: catColor(v.category), display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, fontWeight: 800, color: "#fff", marginBottom: 3 }}>
                {v.name.split(" ").map(n => n[0]).join("").slice(0,2)}
              </div>
            </div>
            <div style={{ fontSize: 8.5, fontWeight: 700, color: "#111", textAlign: "center", lineHeight: 1.2 }}>{v.name.split(" ")[0]}</div>
            <div style={{ fontSize: 7.5, color: "#6B7280", textAlign: "center", lineHeight: 1.2, marginTop: 1 }}>{v.specialty || v.category}</div>
          </>
        )}
        {isVisitor && !v && <div style={{ fontSize: 9, color: "#9CA3AF", textAlign: "center" }}>{isOver ? "Drop here" : "Empty"}</div>}
        {!isVisitor && m && (
          <>
            <div style={{ width: 26, height: 26, borderRadius: "50%", background: catColor(m.category), display: "flex", alignItems: "center", justifyContent: "center", fontSize: 10, fontWeight: 800, color: "#fff", marginBottom: 3 }}>
              {m.name.split(" ").map(n => n[0]).join("").slice(0,2)}
            </div>
            <div style={{ fontSize: 8, fontWeight: 700, color: "#1B2A4A", textAlign: "center", lineHeight: 1.2 }}>{m.name.split(" ")[0]}</div>
            <div style={{ fontSize: 7, color: "#6B7280", textAlign: "center", lineHeight: 1.2, marginTop: 1 }}>{m.specialty?.split(" ").slice(0,2).join(" ")}</div>
          </>
        )}
      </div>
    );
  };

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
        <div>
          <div style={{ fontWeight: 800, fontSize: 15, color: "#111" }}>🪑 Strategic Seat Planner</div>
          <div style={{ fontSize: 11, color: "#6B7280", marginTop: 2 }}>Drag visitors to swap seats. Inner U = visitor row (6 seats). Outer = member seats (10).</div>
        </div>
        <button onClick={autoAssign} style={{ background: "linear-gradient(135deg, #1B2A4A, #8B1A1A)", color: "#fff", border: "none", borderRadius: 8, padding: "8px 16px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>⚡ Auto-Assign by Affinity</button>
      </div>
      <Card style={{ padding: 20, background: "#F8FAFC", position: "relative", overflow: "hidden" }}>
        <div style={{ textAlign: "center", fontSize: 10, fontWeight: 700, color: "#9CA3AF", textTransform: "uppercase", letterSpacing: 2, marginBottom: 14 }}>Conference Room — BNI Insomniacs</div>
        <div style={{ display: "flex", justifyContent: "center", marginBottom: 8 }}>
          <div style={{ background: "#1B2A4A", color: "#fff", borderRadius: 8, padding: "6px 32px", fontSize: 10, fontWeight: 700, letterSpacing: 1 }}>📋 PRESENTER / EDUCATION CHAIR</div>
        </div>
        <div style={{ position: "relative", margin: "0 auto", maxWidth: 680 }}>
          <div style={{ display: "flex", gap: 6, justifyContent: "space-between", marginBottom: 6 }}>
            {[0,1,2,3,4].map(i => <SeatChip key={`M${i+1}`} seatKey={`M${i+1}`} isVisitor={false} />)}
          </div>
          <div style={{ background: "linear-gradient(135deg, #1B2A4A 0%, #243555 100%)", borderRadius: 12, padding: "10px 20px", margin: "0 4px", minHeight: 70, display: "flex", alignItems: "center", justifyContent: "center", boxShadow: "0 4px 20px rgba(27,42,74,0.3)", marginBottom: 6 }}>
            <div style={{ color: "rgba(255,255,255,0.15)", fontSize: 13, fontWeight: 700, letterSpacing: 4, textTransform: "uppercase" }}>◈ CONFERENCE TABLE ◈</div>
          </div>
          <div style={{ background: "rgba(139,26,26,0.06)", borderRadius: 10, padding: "8px 4px", border: "1.5px dashed rgba(139,26,26,0.25)", marginBottom: 6 }}>
            <div style={{ fontSize: 8.5, fontWeight: 700, color: "#8B1A1A", textAlign: "center", textTransform: "uppercase", letterSpacing: 1, marginBottom: 6 }}>★ Visitor Row — Inner U (drag to rearrange)</div>
            <div style={{ display: "flex", gap: 6, justifyContent: "space-between" }}>
              {visitorSeats.map(sKey => <SeatChip key={sKey} seatKey={sKey} isVisitor={true} />)}
            </div>
          </div>
          <div style={{ display: "flex", gap: 6, justifyContent: "space-between" }}>
            {[5,6,7,8,9].map(i => <SeatChip key={`M${i+1}`} seatKey={`M${i+1}`} isVisitor={false} />)}
          </div>
        </div>
        <div style={{ textAlign: "center", marginTop: 12 }}>
          <div style={{ display: "inline-flex", alignItems: "center", gap: 8, background: "#F1F5F9", borderRadius: 20, padding: "4px 16px", fontSize: 10, color: "#6B7280", fontWeight: 600 }}>↑ ENTRANCE / REGISTRATION TABLE ↑</div>
        </div>
      </Card>

      {unassigned.length > 0 && (
        <Card style={{ marginTop: 12, background: "#FFF7ED", borderColor: "#F59E0B" }}>
          <div style={{ fontSize: 11, fontWeight: 700, color: "#92400E", marginBottom: 8 }}>⚠️ Unassigned Visitors ({unassigned.length})</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
            {unassigned.map(v => (
              <div key={v.id} draggable onDragStart={() => setDragging(v.id)}
                style={{ background: catColor(v.category), color: "#fff", borderRadius: 8, padding: "5px 10px", fontSize: 11, fontWeight: 600, cursor: "grab", display: "flex", alignItems: "center", gap: 6 }}>
                <div style={{ width: 20, height: 20, borderRadius: "50%", background: "rgba(255,255,255,0.25)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 9, fontWeight: 800 }}>
                  {v.name.split(" ").map(n => n[0]).join("").slice(0,2)}
                </div>
                {v.name.split(" ")[0]}
              </div>
            ))}
          </div>
        </Card>
      )}
    </div>
  );
}

// ═══════════════════════════════════════════
// PRINTABLE LIST  (uses members from state now)
// ═══════════════════════════════════════════
function PrintBox({ label }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 4 }}>
      <div style={{ width: 13, height: 13, border: "1.5px solid #374151", borderRadius: 2, flexShrink: 0, background: "#fff", display: "inline-block" }} />
      {label && <span style={{ fontSize: 9, color: "#374151" }}>{label}</span>}
    </div>
  );
}

function PrintableVisitorList({ visitors, meetingDate, asks, members, aiMatches, openCategories }) {
  const formatted = (() => {
    const d = new Date(meetingDate + "T00:00:00");
    return d.toLocaleDateString("en-GB", { weekday: "long", day: "numeric", month: "long", year: "numeric" });
  })();
  const getTopMatches = (v) => findMatches(v, asks, members).filter(m => m.score >= 50).slice(0, 2);
  const WALK_IN_ROWS = 3;
  const SUBSTITUTE_ROWS = 8;
  const TH = { padding: "8px 8px", textAlign: "left", fontWeight: 800, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.3 };

  return (
    <div style={{ fontFamily: "'Arial', sans-serif", background: "#fff", color: "#111", padding: "18px 24px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", borderBottom: "3px solid #8B1A1A", paddingBottom: 10, marginBottom: 14 }}>
        <div>
          <div style={{ fontSize: 22, fontWeight: 900, color: "#8B1A1A", letterSpacing: -0.5 }}>BNI Insomniacs</div>
          <div style={{ fontSize: 13, color: "#374151", marginTop: 2 }}>Visitor Host Command Centre</div>
        </div>
        <div style={{ textAlign: "right" }}>
          <div style={{ fontSize: 14, fontWeight: 800, color: "#1B2A4A" }}>Weekly Visitor List</div>
          <div style={{ fontSize: 13, color: "#6B7280", marginTop: 2 }}>{formatted}</div>
          <div style={{ fontSize: 11, color: "#9CA3AF", marginTop: 2 }}>
            {visitors.length} visitor{visitors.length !== 1 ? "s" : ""} registered &nbsp;•&nbsp;
            Printed: {new Date().toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" })}
          </div>
        </div>
      </div>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
        <thead>
          <tr style={{ background: "#1B2A4A", color: "#fff" }}>
            <th style={{ ...TH, width: 24 }}>#</th>
            <th style={{ ...TH, width: 48 }}>Arr.</th>
            <th style={{ ...TH, width: 40 }}>Paid</th>
            <th style={{ ...TH, width: "21%" }}>Visitor Name</th>
            <th style={{ ...TH, width: "23%" }}>Business</th>
            <th style={{ ...TH, width: "20%" }}>Category</th>
            <th style={{ ...TH, width: "15%" }}>Invited By</th>
            <th style={{ ...TH }}>Phone</th>
          </tr>
        </thead>
        <tbody>
          {visitors.map((v, i) => {
            const ai = aiMatches?.[String(v.id)] || [];
            const topMatches = ai.length > 0 ? [] : getTopMatches(v);
            const hasMatches = ai.length > 0 || topMatches.length > 0;
            const rowBg = i % 2 === 0 ? "#fff" : "#F7F8FB";
            return (
              <Fragment key={v.id}>
                <tr style={{ background: rowBg, borderBottom: hasMatches ? "none" : "1px solid #D1D5DB", verticalAlign: "top" }}>
                  <td style={{ padding: "9px 8px 6px", color: "#9CA3AF", fontWeight: 700, fontSize: 12 }}>{i + 1}</td>
                  <td style={{ padding: "9px 8px 6px" }}><PrintBox /></td>
                  <td style={{ padding: "9px 8px 6px" }}><PrintBox /></td>
                  <td style={{ padding: "9px 8px 6px", fontWeight: 800, color: "#111", fontSize: 13 }}>
                    {v.name}
                    <div style={{ fontSize: 10, color: "#6B7280", fontWeight: 400, marginTop: 2 }}>{STATUS_COLORS[v.status]?.label || v.status}</div>
                  </td>
                  <td style={{ padding: "9px 8px 6px", color: "#374151", fontSize: 12 }}>{v.business}</td>
                  <td style={{ padding: "9px 8px 6px" }}>
                    <div style={{ fontWeight: 700, color: "#4338CA", fontSize: 12 }}>{v.category || "—"}</div>
                    <div style={{ color: "#6B7280", fontSize: 10, marginTop: 1 }}>{v.specialty || ""}</div>
                  </td>
                  <td style={{ padding: "9px 8px 6px", color: "#374151", fontSize: 12 }}>{v.invitedBy || "—"}</td>
                  <td style={{ padding: "9px 8px 6px", color: "#374151", fontSize: 12 }}>{v.phone || "—"}</td>
                </tr>
                {hasMatches && (
                  <tr style={{ background: rowBg, borderBottom: "1px solid #D1D5DB" }}>
                    <td colSpan={3} style={{ padding: "0 8px 8px" }}>
                      <div style={{ fontSize: 9, fontWeight: 800, color: "#7C3AED", textTransform: "uppercase", letterSpacing: 0.5, textAlign: "right" }}>Introduce&nbsp;to&nbsp;→</div>
                    </td>
                    <td colSpan={5} style={{ padding: "0 8px 8px" }}>
                      <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                        {ai.length > 0 ? ai.map((m, mi) => {
                          const mem = members.find(mm => mm.name.toLowerCase() === (m.memberName || "").toLowerCase());
                          return (
                            <div key={mi} style={{ display: "inline-flex", alignItems: "center", gap: 6, border: "1px solid #DDD6FE", borderRadius: 6, padding: "3px 9px", background: "#fff" }}>
                              <div style={{ width: 8, height: 8, borderRadius: "50%", flexShrink: 0, background: m.source === "ask" ? "#EF4444" : "#7C3AED" }} />
                              <span style={{ fontWeight: 800, fontSize: 12, color: "#111", whiteSpace: "nowrap" }}>{m.memberName}</span>
                              <span style={{ fontSize: 10, color: "#6B7280", whiteSpace: "nowrap" }}>{m.source === "ask" ? "★ Ask" : "🤖 Synergy"}</span>
                              {mem && <span style={{ fontSize: 10, color: "#9CA3AF", whiteSpace: "nowrap" }}>({mem.specialty})</span>}
                              {m.reason && <span style={{ fontSize: 10.5, color: "#4B5563", fontStyle: "italic" }}>— {m.reason}</span>}
                            </div>
                          );
                        }) : topMatches.map((m, mi) => (
                          <div key={mi} style={{ display: "inline-flex", alignItems: "center", gap: 6, border: "1px solid #E5E7EB", borderRadius: 6, padding: "3px 9px", background: "#fff" }}>
                            <div style={{ width: 8, height: 8, borderRadius: "50%", flexShrink: 0, background: m.score >= 90 ? "#EF4444" : m.score >= 70 ? "#F59E0B" : "#3B82F6" }} />
                            <span style={{ fontWeight: 800, fontSize: 12, color: "#111", whiteSpace: "nowrap" }}>{m.member?.name}</span>
                            <span style={{ fontSize: 10, color: "#6B7280", whiteSpace: "nowrap" }}>{m.type === "ask" ? "★ Ask" : "Contact Sphere"}</span>
                            <span style={{ fontSize: 10, color: "#9CA3AF", whiteSpace: "nowrap" }}>({m.member?.specialty})</span>
                          </div>
                        ))}
                      </div>
                    </td>
                  </tr>
                )}
              </Fragment>
            );
          })}
          {Array.from({ length: WALK_IN_ROWS }).map((_, i) => (
            <tr key={"walkin-" + i} style={{ background: i % 2 === 0 ? "#FFFBF0" : "#FFF8E8", borderBottom: "1px solid #FDE68A" }}>
              <td style={{ padding: "12px 8px", color: "#D97706", fontWeight: 700, fontSize: 12 }}>{visitors.length + i + 1}</td>
              <td style={{ padding: "12px 8px" }}><PrintBox /></td>
              <td style={{ padding: "12px 8px" }}><PrintBox /></td>
              <td style={{ padding: "12px 8px", borderBottom: "1px dotted #FCD34D" }}>
                <div style={{ fontSize: 9, color: "#D97706", fontWeight: 700, marginBottom: 2 }}>WALK-IN</div>
                <div style={{ height: 18 }} />
              </td>
              <td style={{ padding: "12px 8px", borderBottom: "1px dotted #FCD34D" }}>&nbsp;</td>
              <td style={{ padding: "12px 8px", borderBottom: "1px dotted #FCD34D" }}>&nbsp;</td>
              <td style={{ padding: "12px 8px", borderBottom: "1px dotted #FCD34D" }}>&nbsp;</td>
              <td style={{ padding: "12px 8px", borderBottom: "1px dotted #FCD34D" }}>&nbsp;</td>
            </tr>
          ))}
        </tbody>
      </table>
      {openCategories?.categories?.length > 0 && (
        <div style={{ marginTop: 18, pageBreakInside: "avoid" }}>
          <div style={{ background: "#8B1A1A", color: "#fff", padding: "8px 12px", fontSize: 12, fontWeight: 800, textTransform: "uppercase", letterSpacing: 0.5, borderRadius: "4px 4px 0 0", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span>🎯 Open Categories — Top 10 Invite Targets</span>
            {openCategories.checkedAt && <span style={{ fontSize: 10, fontWeight: 600, opacity: 0.85, textTransform: "none", letterSpacing: 0 }}>Checked: {new Date(openCategories.checkedAt).toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" })}</span>}
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", border: "1px solid #D1D5DB", borderTop: "none" }}>
            {openCategories.categories.slice(0, 10).map((c, i) => (
              <div key={i} style={{ display: "flex", gap: 8, padding: "7px 10px", borderBottom: i < openCategories.categories.slice(0, 10).length - 2 ? "1px solid #E5E7EB" : "none", borderRight: i % 2 === 0 ? "1px solid #E5E7EB" : "none", background: c.fit === "high" ? "#FFFBEB" : "#fff" }}>
                <div style={{ width: 18, height: 18, borderRadius: "50%", background: c.fit === "high" ? "#8B1A1A" : "#6B7280", color: "#fff", fontSize: 10, fontWeight: 800, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, marginTop: 1 }}>{i + 1}</div>
                <div style={{ minWidth: 0 }}>
                  <span style={{ fontWeight: 800, fontSize: 12, color: "#111" }}>{c.category}</span>
                  {c.fit === "high" && <span style={{ fontSize: 9, fontWeight: 800, color: "#B45309", marginLeft: 6 }}>★ TOP DEMAND</span>}
                  {c.chapters > 0 && <span style={{ fontSize: 9, fontWeight: 700, color: "#6D28D9", marginLeft: 6 }}>📊 {c.chapters} UAE ch.</span>}
                  {c.synergyWith && <div style={{ fontSize: 10, color: "#4338CA", marginTop: 1 }}>↔ {c.synergyWith}</div>}
                  {c.reason && <div style={{ fontSize: 10.5, color: "#6B7280", fontStyle: "italic", marginTop: 1 }}>{c.reason}</div>}
                </div>
              </div>
            ))}
          </div>
        </div>
      )}
      <div style={{ marginTop: 18, pageBreakInside: "avoid" }}>
        <div style={{ background: "#1B2A4A", color: "#fff", padding: "8px 12px", fontSize: 12, fontWeight: 800, textTransform: "uppercase", letterSpacing: 0.5, borderRadius: "4px 4px 0 0" }}>Member Substitutes</div>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
          <thead>
            <tr style={{ background: "#F1F5F9" }}>
              {["#","BNI Member Name (Absent)","Substitute Name","Substitute Business / Company","Phone","Arrived","Paid"].map((h, i) => (
                <th key={i} style={{ padding: "7px 9px", textAlign: "left", fontWeight: 700, fontSize: 11, textTransform: "uppercase", letterSpacing: 0.3, color: "#374151", border: "1px solid #D1D5DB" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {Array.from({ length: SUBSTITUTE_ROWS }).map((_, i) => (
              <tr key={i} style={{ background: i % 2 === 0 ? "#fff" : "#F8F9FC" }}>
                <td style={{ padding: "11px 9px", border: "1px solid #E5E7EB", color: "#9CA3AF", fontWeight: 700, textAlign: "center", fontSize: 12 }}>{i + 1}</td>
                <td style={{ padding: "11px 9px", border: "1px solid #E5E7EB" }}>&nbsp;</td>
                <td style={{ padding: "11px 9px", border: "1px solid #E5E7EB" }}>&nbsp;</td>
                <td style={{ padding: "11px 9px", border: "1px solid #E5E7EB" }}>&nbsp;</td>
                <td style={{ padding: "11px 9px", border: "1px solid #E5E7EB" }}>&nbsp;</td>
                <td style={{ padding: "11px 9px", border: "1px solid #E5E7EB", textAlign: "center" }}><PrintBox /></td>
                <td style={{ padding: "11px 9px", border: "1px solid #E5E7EB", textAlign: "center" }}><PrintBox /></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div style={{ marginTop: 14, paddingTop: 6, borderTop: "2px solid #8B1A1A", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div style={{ fontSize: 10, color: "#9CA3AF" }}>BNI Insomniacs • Visitor Host Programme • Confidential</div>
        <div style={{ fontSize: 10, color: "#9CA3AF" }}>Givers Gain®</div>
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════
// TABS
// ═══════════════════════════════════════════
const TABS = [
  { label: "Dashboard", icon: "📊" },
  { label: "Visitors", icon: "👥" },
  { label: "Asks", icon: "🎯" },
  { label: "Connect", icon: "🔗" },
  { label: "AI Match", icon: "🤖" },
  { label: "Seat Planner", icon: "🪑" },
  { label: "Members", icon: "📇" },
  { label: "Archive", icon: "🗄️" },
  { label: "Follow-Up", icon: "📬" },
];

function DashboardTab({ visitors, asks, members, archived }) {
  const thisWeek = visitors.filter(v => v.date === MEETING_DATE);
  const attended = visitors.filter(v => ["attended","oriented","applied","joined"].includes(v.status)).length;
  const applied = visitors.filter(v => ["applied","joined"].includes(v.status)).length;
  const ratio = attended > 0 ? Math.round((applied / attended) * 100) : 0;
  const catCounts = {};
  members.forEach(m => { catCounts[m.category] = (catCounts[m.category] || 0) + 1; });

  return <div>
    <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 16 }}>
      <StatCard label="This Week" value={thisWeek.length} sub="Visitors registered" color="#4338CA" />
      <StatCard label="Members" value={members.length} sub="BNI Insomniacs" color="#065F46" />
      <StatCard label="Open Asks" value={asks.filter(isActiveAsk).length} sub="Active (last 6 weeks)" color="#D97706" />
      <StatCard label="Closing Ratio" value={`${ratio}%`} sub={`${applied}/${attended}`} color="#9D174D" />
      <StatCard label="Archived" value={archived.length} sub="Past visitors" color="#6B21A8" />
    </div>
    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
      <Card>
        <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 10 }}>📋 7-Touchpoint Checklist</div>
        {["A. Invitation — Members registered visitors","B. Registration — Confirmations sent","C. Pre-Meeting Call — All visitors called","D. The Welcome — Greeter + Registration ready","E. Meeting Experience — Orientation Host briefed","F. Visitor Interest — VO slides, ask 'Would you like to join?'","G. Follow-Up — Call within 2 hrs, 3-Response ready"].map((item, i) => (
          <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, padding: "4px 0", fontSize: 12, color: "#374151" }}>
            <div style={{ width: 16, height: 16, borderRadius: 4, border: "2px solid #D1D5DB", flexShrink: 0 }} />
            {item}
          </div>
        ))}
      </Card>
      <Card>
        <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 10 }}>🏷️ Chapter Categories ({Object.keys(catCounts).length})</div>
        <div style={{ maxHeight: 220, overflowY: "auto" }}>
          {Object.entries(catCounts).sort((a,b) => b[1]-a[1]).map(([cat, count]) => (
            <div key={cat} style={{ display: "flex", justifyContent: "space-between", padding: "3px 0", fontSize: 12, borderBottom: "1px solid #F3F4F6" }}>
              <span style={{ color: "#374151" }}>{cat}</span>
              <span style={{ fontWeight: 700, color: "#6B7280" }}>{count}</span>
            </div>
          ))}
        </div>
      </Card>
    </div>
  </div>;
}

// ═══════════════════════════════════════════
// VISITORS TAB — with Edit, Delete, Archive, Validate
// ═══════════════════════════════════════════
function VisitorsTab({ visitors, setVisitors, asks, members, archived, setArchived }) {
  const [showForm, setShowForm] = useState(false);
  const [viewMode, setViewMode] = useState("list");
  const [printDate, setPrintDate] = useState(MEETING_DATE);
  const [expandedId, setExpandedId] = useState(null);
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState({ name: "", business: "", phone: "", email: "", invitedBy: "", category: "", specialty: "", date: MEETING_DATE });
  const [editForm, setEditForm] = useState({});
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [confirmArchive, setConfirmArchive] = useState(null);
  const [showPrintModal, setShowPrintModal] = useState(false);
  const [activePanel, setActivePanel] = useState({});  // { [visitorId]: 'brief' | 'validate' }

  // ── Print View AI-Match: { [visitorId]: [{ memberName, reason, source }] } ──
  const [printMatches, setPrintMatches] = useState({});
  const [matchLoading, setMatchLoading] = useState(false);
  const [matchError, setMatchError] = useState("");

  // ── Open Categories: { checkedAt, categories: [{ category, synergyWith, reason, fit }] } ──
  const OPEN_CATS_KEY = "bni-open-categories";
  const [openCats, setOpenCats] = useState(null);
  const [showOpenCats, setShowOpenCats] = useState(false);
  const [catsLoading, setCatsLoading] = useState(false);
  const [catsError, setCatsError] = useState("");

  // Restore last weekly scan from this browser (survives page refreshes, unlike AI matches)
  useEffect(() => {
    try {
      const raw = localStorage.getItem(OPEN_CATS_KEY);
      if (raw) setOpenCats(JSON.parse(raw));
    } catch { /* corrupt storage — ignore, user can re-run */ }
  }, []);

  const daysSinceCatCheck = openCats?.checkedAt ? Math.floor((Date.now() - new Date(openCats.checkedAt).getTime()) / 86400000) : null;

  const allCategoriesPresent = [...new Set(members.map(m => m.category))].sort();

  useEffect(() => {
    if (!document.getElementById("bni-spin-style")) {
      const style = document.createElement("style");
      style.id = "bni-spin-style";
      style.textContent = "@keyframes spin { to { transform: rotate(360deg); } }";
      document.head.appendChild(style);
    }
  }, []);

  const allDates = [...new Set(visitors.map(v => v.date))].sort().reverse();
  const printVisitors = visitors.filter(v => v.date === printDate);

  // ═══════════════════════════════════════════
  // PRINT VIEW AI-MATCH — fills "BNI Matches — Introduce To" using asks + full member list
  // ═══════════════════════════════════════════
  const runPrintAIMatch = async () => {
    if (printVisitors.length === 0) return;
    setMatchLoading(true);
    setMatchError("");
    try {
      const openAsks = asks.filter(a => a.status === "open");
      const visitorLines = printVisitors.map(v =>
        `- id:${v.id} | ${v.name} | Business: ${v.business || "?"} | Category: ${v.category || "?"} | Specialty: ${v.specialty || "?"}`
      ).join("\n");
      const askLines = openAsks.length === 0 ? "(none)" : openAsks.map(a =>
        `- ${a.memberName} is looking for: ${[a.targetName, a.targetCompany, a.targetRole, a.targetCategory].filter(Boolean).join(" / ") || "?"}${a.notes ? ` — "${a.notes}"` : ""}`
      ).join("\n");
      const memberLines = members.map(m => `- ${m.name} | ${m.category} | ${m.specialty}`).join("\n");

      const prompt = `You are the introduction-matching engine for BNI Insomniacs, a BNI chapter in Dubai. For each visitor attending this week's meeting, recommend 1-3 chapter members they should be introduced to.

VISITORS THIS WEEK:
${visitorLines}

OPEN MEMBER ASKS (highest priority — if a visitor could fulfil an ask, always match them):
${askLines}

CHAPTER MEMBERS (only recommend names EXACTLY as written here):
${memberLines}

MATCHING RULES:
1. First priority: visitors who could fulfil an open ask → source "ask".
2. Second priority: strong business synergy — complementary services, same client base, contact-sphere overlap, likely referral partners → source "synergy".
3. Every visitor must get at least 1 match; give 2-3 where genuinely useful. Never invent member names.
4. Do NOT match a visitor to the member who invited them.
5. Keep each reason under 12 words, specific and actionable.

Respond with ONLY valid JSON, no markdown, no preamble:
{"matches":[{"visitorId":"<id exactly as given>","introduceTo":[{"memberName":"<exact member name>","reason":"<short reason>","source":"ask" or "synergy"}]}]}`;

      const result = await callClaude(prompt, 3000);
      const memberNamesLower = members.map(m => m.name.toLowerCase());
      const map = {};
      (result.matches || []).forEach(m => {
        const valid = (m.introduceTo || []).filter(x => x.memberName && memberNamesLower.includes(x.memberName.toLowerCase())).slice(0, 3);
        if (valid.length) map[String(m.visitorId)] = valid;
      });
      setPrintMatches(prev => ({ ...prev, ...map }));
      if (Object.keys(map).length === 0) setMatchError("AI returned no usable matches — try again.");
    } catch (e) {
      console.error("Print AI-Match failed:", e);
      setMatchError(e.message || "AI match failed — check the API key and model string.");
    }
    setMatchLoading(false);
  };

  // ═══════════════════════════════════════════
  // OPEN CATEGORIES SCAN — data-driven, from the UAE consolidated member list.
  // Shortlists categories filled elsewhere in the UAE but open in Insomniacs,
  // cross-checked against the LIVE member roster, ranked by nationwide demand.
  // AI is only used to annotate referral synergy — if that call fails, the
  // data-driven list still stands (the scan itself can never come back empty).
  // Run weekly via the button; result is saved in this browser (localStorage).
  // ═══════════════════════════════════════════
  const normCat = (s) => (s || "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();

  const runOpenCategoryScan = async () => {
    setCatsLoading(true);
    setCatsError("");
    try {
      // 1. Cross-check the reference list against the live roster (category + specialty)
      const taken = members.flatMap(m => [normCat(m.category), normCat(m.specialty)]).filter(Boolean);
      const stillOpen = UAE_OPEN_CATEGORY_POOL.filter(([cat]) => {
        const n = normCat(cat);
        return !taken.some(t => t === n || (t.length >= 12 && n.length >= 12 && (t.includes(n) || n.includes(t))));
      });
      if (stillOpen.length === 0) throw new Error("Every category in the UAE reference list is now taken — the reference data needs a refresh.");

      // 2. Top 10 by demand (chapters filled in, then members nationwide) — already pre-sorted
      const top = stillOpen.slice(0, 10).map(([category, group, chapters, memberCount], i) => ({
        category, group, chapters, memberCount,
        fit: i < 5 ? "high" : "good",
        synergyWith: "",
        reason: `Filled in ${chapters} other UAE chapter${chapters === 1 ? "" : "s"} (${memberCount} member${memberCount === 1 ? "" : "s"} nationwide)`,
      }));

      // 3. Optional AI pass — annotate each pick with referral synergy. Failure is non-fatal.
      try {
        const memberLines = members.map(m => `- ${m.category}${m.specialty ? ` (${m.specialty})` : ""}`).join("\n");
        const pickLines = top.map(c => `- ${c.category} [group: ${c.group}]`).join("\n");
        const prompt = `You are the membership-growth advisor for BNI Insomniacs, a BNI chapter in Dubai. The 10 categories below are confirmed OPEN in this chapter (they are filled in other UAE chapters). Do NOT add, remove, rename, or re-rank them — only annotate each one.

CURRENT CHAPTER CATEGORIES:
${memberLines}

OPEN CATEGORIES TO ANNOTATE:
${pickLines}

For each open category provide:
- "synergyWith": 2-3 CURRENT chapter categories it would trade referrals with, comma-separated, exactly as written in the current list.
- "reason": under 14 words on why it fits THIS chapter's referral network.

Respond with ONLY valid JSON, no markdown, no preamble:
{"categories":[{"category":"<exactly as given>","synergyWith":"...","reason":"..."}]}`;
        const result = await callClaude(prompt, 2000);
        const notes = {};
        (result.categories || []).forEach(c => { if (c.category) notes[normCat(c.category)] = c; });
        top.forEach(t => {
          const n = notes[normCat(t.category)];
          if (n) {
            if (n.synergyWith) t.synergyWith = String(n.synergyWith);
            if (n.reason) t.reason = String(n.reason);
          }
        });
      } catch (aiErr) {
        console.warn("Synergy annotation skipped (data-driven list kept):", aiErr);
      }

      const payload = { checkedAt: new Date().toISOString(), categories: top };
      setOpenCats(payload);
      try { localStorage.setItem(OPEN_CATS_KEY, JSON.stringify(payload)); } catch { /* storage full/blocked — session only */ }
    } catch (e) {
      console.error("Open category scan failed:", e);
      setCatsError(e.message || "Scan failed.");
    }
    setCatsLoading(false);
  };

  const addVisitor = async () => {
    if (!form.name || !form.business) return;
    const { data } = await supabase.from("visitors").insert([{
      name: form.name, business: form.business, phone: form.phone,
      email: form.email, invited_by: form.invitedBy, category: form.category,
      specialty: form.specialty, date: form.date, status: "registered",
      call_notes: "", seat_assignment: "", follow_up_response: null, bio: null
    }]).select();
    if (data?.[0]) setVisitors(p => [...p, { ...data[0], invitedBy: data[0].invited_by, callNotes: data[0].call_notes, seatAssignment: data[0].seat_assignment, followUpResponse: data[0].follow_up_response }]);
    setForm({ name: "", business: "", phone: "", email: "", invitedBy: "", category: "", specialty: "", date: MEETING_DATE });
    setShowForm(false);
  };

  const updateStatus = async (id, status) => {
    await supabase.from("visitors").update({ status }).eq("id", id);
    setVisitors(p => p.map(v => v.id === id ? { ...v, status } : v));
  };
  const saveBio = async (id, bio) => {
    await supabase.from("visitors").update({ bio }).eq("id", id);
    setVisitors(p => p.map(v => v.id === id ? { ...v, bio } : v));
  };
  const saveValidation = (id, validation) => setVisitors(p => p.map(v => v.id === id ? { ...v, validation } : v));

  const startEdit = (v) => { setEditingId(v.id); setEditForm({ ...v }); setExpandedId(null); };
  const cancelEdit = () => { setEditingId(null); setEditForm({}); };
  const saveEdit = async () => {
    await supabase.from("visitors").update({
      name: editForm.name, business: editForm.business, phone: editForm.phone,
      email: editForm.email, invited_by: editForm.invitedBy, category: editForm.category,
      specialty: editForm.specialty, date: editForm.date, status: editForm.status,
      call_notes: editForm.callNotes, seat_assignment: editForm.seatAssignment,
    }).eq("id", editingId);
    setVisitors(p => p.map(v => v.id === editingId ? { ...v, ...editForm } : v));
    setEditingId(null);
    setEditForm({});
  };

  const requestDelete = (v) => setConfirmDelete(v);
  const doDelete = async () => {
    await supabase.from("visitors").delete().eq("id", confirmDelete.id);
    setVisitors(p => p.filter(v => v.id !== confirmDelete.id));
    setConfirmDelete(null);
  };

  const requestArchive = (v) => setConfirmArchive(v);
  const doArchive = () => {
    const archivedAt = new Date().toISOString();
    setArchived(p => [...p, { ...confirmArchive, archivedAt }]);
    setVisitors(p => p.filter(v => v.id !== confirmArchive.id));
    setConfirmArchive(null);
  };

  // ═══════════════════════════════════════════
  // EXCEL BULK IMPORT / EXPORT
  // ═══════════════════════════════════════════
  const [importPreview, setImportPreview] = useState(null); // array of parsed rows or null
  const [importing, setImporting] = useState(false);
  const fileInputRef = useRef(null);

  const TEMPLATE_ROWS = 150; // dropdowns are applied to this many entry rows

  const downloadTemplate = async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Visitors");
    const lists = wb.addWorksheet("Lists");

    // ── Lists sheet feeds the dropdowns (also handy as a reference) ──
    const memberNames = members.map(m => m.name).sort((a, b) => a.localeCompare(b));
    const categories = [...new Set([...members.map(m => m.category), ...ALL_BNI_CATEGORIES])].sort();
    const specialties = [...new Set(members.map(m => m.specialty))].sort();
    const dates = upcomingWednesdays(8);
    lists.getRow(1).values = ["Members", "Categories", "Specialties", "Meeting Dates"];
    lists.getRow(1).font = { bold: true };
    memberNames.forEach((v, i) => { lists.getCell(`A${i + 2}`).value = v; });
    categories.forEach((v, i) => { lists.getCell(`B${i + 2}`).value = v; });
    specialties.forEach((v, i) => { lists.getCell(`C${i + 2}`).value = v; });
    dates.forEach((v, i) => { lists.getCell(`D${i + 2}`).value = v; });
    lists.columns = [{ width: 28 }, { width: 28 }, { width: 36 }, { width: 16 }];

    // ── Visitors sheet: type only Name/Business/Phone/Email; rest are dropdowns ──
    ws.columns = [
      { header: "Name", width: 22 }, { header: "Business", width: 26 },
      { header: "Phone", width: 16 }, { header: "Email", width: 28 },
      { header: "Invited By", width: 26 }, { header: "Category", width: 28 },
      { header: "Specialty", width: 30 }, { header: "Meeting Date", width: 16 },
    ];
    const head = ws.getRow(1);
    head.font = { bold: true, color: { argb: "FFFFFFFF" } };
    head.eachCell(c => { c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF8B1A1A" } }; });
    ws.views = [{ state: "frozen", ySplit: 1 }];

    const strictList = (range, what) => ({
      type: "list", allowBlank: true, formulae: [range],
      showErrorMessage: true, errorStyle: "stop",
      errorTitle: `Invalid ${what}`, error: `Please pick a ${what} from the dropdown — typing is disabled to avoid errors.`,
    });
    for (let r = 2; r <= TEMPLATE_ROWS + 1; r++) {
      ws.getCell(`E${r}`).dataValidation = strictList(`Lists!$A$2:$A$${memberNames.length + 1}`, "member");
      ws.getCell(`F${r}`).dataValidation = strictList(`Lists!$B$2:$B$${categories.length + 1}`, "category");
      ws.getCell(`G${r}`).dataValidation = {
        type: "list", allowBlank: true, formulae: [`Lists!$C$2:$C$${specialties.length + 1}`],
        showErrorMessage: true, errorStyle: "warning",
        errorTitle: "New specialty", error: "This specialty isn't in the standard list. Click Yes to keep it anyway.",
      };
      ws.getCell(`H${r}`).dataValidation = strictList(`Lists!$D$2:$D$${dates.length + 1}`, "meeting date");
    }

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `BNI_Visitor_Import_Template_${MEETING_DATE}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const exportVisitors = () => {
    const rows = visitors.map(v => ({
      Name: v.name, Business: v.business, Phone: v.phone, Email: v.email,
      "Invited By": v.invitedBy || "", Category: v.category || "", Specialty: v.specialty || "",
      "Meeting Date": v.date || "", Status: v.status || "", "Call Notes": v.callNotes || "",
    }));
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = Object.keys(rows[0] || { a: 1 }).map(() => ({ wch: 22 }));
    XLSX.utils.book_append_sheet(wb, ws, "Visitors");
    XLSX.writeFile(wb, `BNI_Visitors_Export_${toYMD(new Date())}.xlsx`);
  };

  const handleImportFile = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { defval: "" });
        const memberNames = members.map(m => m.name.toLowerCase());
        // Header lookup is case/spacing tolerant
        const get = (row, ...keys) => {
          const rowKeys = Object.keys(row);
          for (const k of keys) {
            const found = rowKeys.find(rk => rk.toLowerCase().replace(/[^a-z]/g, "").startsWith(k));
            if (found !== undefined) return row[found];
          }
          return "";
        };
        const parsed = raw.map((row, i) => {
          const name = String(get(row, "name") || "").trim();
          const business = String(get(row, "business", "company") || "").trim();
          const phone = String(get(row, "phone", "mobile") || "").trim();
          const email = String(get(row, "email") || "").trim();
          const invitedBy = String(get(row, "invitedby", "inviter") || "").trim();
          const category = String(get(row, "category") || "").trim();
          const specialty = String(get(row, "specialty", "classification") || "").trim();
          const date = normalizeExcelDate(get(row, "meetingdate", "date"));
          const warnings = [];
          let status = "ok";
          if (!name) { status = "error"; warnings.push("Name is required"); }
          if (!business) { status = "error"; warnings.push("Business is required"); }
          if (invitedBy && !memberNames.includes(invitedBy.toLowerCase())) warnings.push(`Inviter "${invitedBy}" not found in member list`);
          const knownCategories = [...new Set([...members.map(m => m.category), ...ALL_BNI_CATEGORIES])].map(c => c.toLowerCase());
          if (category && !knownCategories.includes(category.toLowerCase())) warnings.push(`"${category}" is not a standard BNI category`);
          if (new Date(date).getDay() !== 3) warnings.push(`${date} is not a Wednesday`);
          const dup = visitors.find(v =>
            (phone && v.phone && v.phone.replace(/\s/g, "") === phone.replace(/\s/g, "")) ||
            (email && v.email && v.email.toLowerCase() === email.toLowerCase()) ||
            (name && v.name.toLowerCase() === name.toLowerCase() && v.business.toLowerCase() === business.toLowerCase())
          );
          if (dup) { status = "duplicate"; warnings.push(`Already in list as "${dup.name}" (${dup.business})`); }
          if (status === "ok" && warnings.length) status = "warning";
          return { rowNum: i + 2, name, business, phone, email, invitedBy, category, specialty, date, status, warnings };
        }).filter(r => r.name || r.business || r.phone || r.email); // skip fully blank rows
        setImportPreview(parsed.length ? parsed : []);
      } catch (err) {
        console.error("Excel parse failed:", err);
        alert("Could not read that file. Please use the downloaded template (.xlsx).");
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = ""; // allow re-selecting the same file
  };

  // ═══════════════════════════════════════════
  // PASTE-TO-ADD — paste the visitor list straight from BNI Connect
  // ═══════════════════════════════════════════
  const [showPaste, setShowPaste] = useState(false);
  const [pasteText, setPasteText] = useState("");
  const [pastePreview, setPastePreview] = useState(null);
  const pasteCategoryOptions = [...new Set([...members.map(m => m.category), ...ALL_BNI_CATEGORIES])].sort();
  const sortedMembers = [...members].sort((a, b) => a.name.localeCompare(b.name));

  const parsePastedList = () => {
    const isDate = (l) => /^\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}$/.test(l);
    const isEmail = (l) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(l);
    const isPhone = (l) => /^[+\d][\d\s\-()]{6,}$/.test(l);
    const lines = pasteText.split(/\r?\n/).map(l => l.trim()).filter(l => l && !/^(edit|delete)$/i.test(l));

    const records = [];
    let cur = { pre: [], date: "", phone: "", email: "" };
    const flush = () => {
      if (cur.pre.length || cur.date || cur.phone || cur.email) records.push(cur);
      cur = { pre: [], date: "", phone: "", email: "" };
    };
    for (const l of lines) {
      if (isDate(l)) { if (cur.date) flush(); cur.date = l; continue; }
      if (isEmail(l)) { cur.email = l; flush(); continue; }           // email closes a record
      if (isPhone(l) && (cur.date || cur.pre.length >= 3)) { if (cur.phone) flush(); cur.phone = l; continue; }
      if (cur.date || cur.phone || cur.email) flush();                // plain text after data = next visitor
      cur.pre.push(l);
    }
    flush();

    const rows = records.filter(r => r.pre.length || r.phone || r.email).map(r => {
      const name = r.pre[0] || "";
      let business = r.pre.length >= 3 ? r.pre[1] : "";
      const specialty = r.pre.length >= 3 ? r.pre.slice(2).join(" ") : (r.pre[1] || "");
      if (/^no company$/i.test(business)) business = "";
      const date = normalizeExcelDate(r.date);
      const warnings = [];
      let status = "ok";
      if (!name) { status = "error"; warnings.push("Could not detect a name"); }
      if (!business) warnings.push("No company");
      if (new Date(date).getDay() !== 3) warnings.push(`${date} is not a Wednesday`);
      const dup = visitors.find(v =>
        (r.phone && v.phone && v.phone.replace(/\D/g, "").slice(-9) === r.phone.replace(/\D/g, "").slice(-9)) ||
        (r.email && v.email && v.email.toLowerCase() === r.email.toLowerCase()) ||
        (name && v.name.toLowerCase() === name.toLowerCase() && (v.business || "").toLowerCase() === business.toLowerCase())
      );
      if (dup) { status = "duplicate"; warnings.push(`Already in list as "${dup.name}"`); }
      if (status === "ok" && warnings.length) status = "warning";
      return { name, business, specialty, phone: r.phone, email: r.email, date, invitedBy: "", category: "", status, warnings };
    });

    if (!rows.length) { alert("Couldn't detect any visitors in that text. Paste the list exactly as copied from BNI Connect."); return; }
    setPastePreview(rows);
    setShowPaste(false);
  };

  const setPasteField = (i, field, val) => setPastePreview(p => p.map((r, idx) => idx === i ? { ...r, [field]: val } : r));
  const setAllInviters = (name) => setPastePreview(p => p.map(r => (r.status === "ok" || r.status === "warning") ? { ...r, invitedBy: name } : r));

  const confirmPasteImport = async () => {
    const toImport = pastePreview.filter(r => r.status === "ok" || r.status === "warning");
    if (!toImport.length) { setPastePreview(null); return; }
    setImporting(true);
    try {
      const { data, error } = await supabase.from("visitors").insert(
        toImport.map(r => ({
          name: r.name, business: r.business, phone: r.phone, email: r.email,
          invited_by: r.invitedBy, category: r.category, specialty: r.specialty,
          date: r.date, status: "registered",
          call_notes: "", seat_assignment: "", follow_up_response: null, bio: null,
        }))
      ).select();
      if (error) throw error;
      if (data?.length) {
        setVisitors(p => [...p, ...data.map(v => ({
          ...v, invitedBy: v.invited_by, callNotes: v.call_notes,
          seatAssignment: v.seat_assignment, followUpResponse: v.follow_up_response,
        }))]);
      }
      setPastePreview(null);
      setPasteText("");
    } catch (err) {
      console.error("Paste import failed:", err);
      alert("Import failed — nothing was saved. Check your connection and try again.");
    }
    setImporting(false);
  };

  const confirmImport = async () => {
    const toImport = importPreview.filter(r => r.status === "ok" || r.status === "warning");
    if (!toImport.length) { setImportPreview(null); return; }
    setImporting(true);
    try {
      const { data, error } = await supabase.from("visitors").insert(
        toImport.map(r => ({
          name: r.name, business: r.business, phone: r.phone, email: r.email,
          invited_by: r.invitedBy, category: r.category, specialty: r.specialty,
          date: r.date, status: "registered",
          call_notes: "", seat_assignment: "", follow_up_response: null, bio: null,
        }))
      ).select();
      if (error) throw error;
      if (data?.length) {
        setVisitors(p => [...p, ...data.map(v => ({
          ...v, invitedBy: v.invited_by, callNotes: v.call_notes,
          seatAssignment: v.seat_assignment, followUpResponse: v.follow_up_response,
        }))]);
      }
      setImportPreview(null);
    } catch (err) {
      console.error("Bulk import failed:", err);
      alert("Import failed — nothing was saved. Check your connection and try again.");
    }
    setImporting(false);
  };

  return <div>
    <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>

    <ConfirmModal
      open={!!confirmDelete}
      title="Delete this visitor?"
      message={confirmDelete ? `This will permanently remove "${confirmDelete.name}" from your active list. This cannot be undone.` : ""}
      confirmLabel="Yes, delete"
      onConfirm={doDelete}
      onCancel={() => setConfirmDelete(null)}
    />
    <ConfirmModal
      open={!!confirmArchive}
      title="Archive this visitor?"
      message={confirmArchive ? `"${confirmArchive.name}" will move from your active list to the Archive tab. You can look them up later by month.` : ""}
      confirmLabel="Yes, archive"
      danger={false}
      onConfirm={doArchive}
      onCancel={() => setConfirmArchive(null)}
    />

    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
      <span style={{ fontWeight: 700, fontSize: 15 }}>All Visitors ({visitors.length})</span>
      <div style={{ display: "flex", gap: 6 }}>
        <div style={{ display: "flex", borderRadius: 8, border: "1px solid #D1D5DB", overflow: "hidden" }}>
          <button onClick={() => setViewMode("list")} style={{ padding: "7px 12px", fontSize: 11, fontWeight: 600, border: "none", cursor: "pointer", background: viewMode === "list" ? "#1B2A4A" : "#fff", color: viewMode === "list" ? "#fff" : "#6B7280" }}>☰ List</button>
          <button onClick={() => setViewMode("print")} style={{ padding: "7px 12px", fontSize: 11, fontWeight: 600, border: "none", cursor: "pointer", background: viewMode === "print" ? "#1B2A4A" : "#fff", color: viewMode === "print" ? "#fff" : "#6B7280" }}>🖨️ Print View</button>
        </div>
        {viewMode === "list" && (
          <>
            <button onClick={() => setShowOpenCats(!showOpenCats)} title="Data scan: top 10 categories filled in other UAE chapters but open in ours" style={{ background: showOpenCats ? "#7C3AED" : "#fff", color: showOpenCats ? "#fff" : "#7C3AED", border: "1px solid #7C3AED", borderRadius: 8, padding: "7px 12px", fontSize: 12, fontWeight: 600, cursor: "pointer", position: "relative" }}>
              🎯 Open Categories
              {daysSinceCatCheck !== null && daysSinceCatCheck >= 7 && <span style={{ position: "absolute", top: -4, right: -4, width: 10, height: 10, borderRadius: "50%", background: "#F59E0B", border: "2px solid #fff" }} />}
            </button>
            <button onClick={downloadTemplate} title="Download a blank Excel template" style={{ background: "#fff", color: "#1B2A4A", border: "1px solid #D1D5DB", borderRadius: 8, padding: "7px 12px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>⬇️ Template</button>
            <button onClick={() => { setShowPaste(!showPaste); setPastePreview(null); }} title="Paste the visitor list copied from BNI Connect" style={{ background: showPaste ? "#1B2A4A" : "#fff", color: showPaste ? "#fff" : "#1B2A4A", border: "1px solid #1B2A4A", borderRadius: 8, padding: "7px 12px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>📋 Paste List</button>
            <button onClick={() => fileInputRef.current?.click()} title="Import visitors from a filled template" style={{ background: "#1B2A4A", color: "#fff", border: "none", borderRadius: 8, padding: "7px 12px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>⬆️ Import Excel</button>
            <button onClick={exportVisitors} title="Export all visitors to Excel" style={{ background: "#fff", color: "#1B2A4A", border: "1px solid #D1D5DB", borderRadius: 8, padding: "7px 12px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>📤 Export</button>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleImportFile} style={{ display: "none" }} />
            <button onClick={() => setShowForm(!showForm)} style={{ background: "#8B1A1A", color: "#fff", border: "none", borderRadius: 8, padding: "7px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>+ Register</button>
          </>
        )}
      </div>
    </div>

    {viewMode === "list" && showOpenCats && (
      <Card style={{ marginBottom: 12, background: "#FAF5FF", borderColor: "#A855F7" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8, marginBottom: 6 }}>
          <div>
            <div style={{ fontSize: 13, fontWeight: 800, color: "#6B21A8" }}>🎯 Open Categories in BNI Insomniacs</div>
            <div style={{ fontSize: 11, color: "#7E22CE", lineHeight: 1.6, marginTop: 2 }}>
              The top 10 invite targets, taken from the consolidated UAE member list (37 chapters, 1,085 members): categories filled elsewhere in the UAE but open in Insomniacs, cross-checked live against our {members.length} members and ranked by nationwide demand. AI adds the referral-synergy notes. Re-run weekly — results are saved in this browser and printed automatically on the visitor sheet.
            </div>
          </div>
          <button onClick={() => setShowOpenCats(false)} style={{ background: "none", border: "none", fontSize: 14, cursor: "pointer", color: "#6B21A8", flexShrink: 0 }}>✕</button>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", marginBottom: openCats ? 10 : 0 }}>
          <button onClick={runOpenCategoryScan} disabled={catsLoading || members.length === 0}
            style={{ background: "#7C3AED", color: "#fff", border: "none", borderRadius: 8, padding: "8px 18px", fontSize: 12, fontWeight: 700, cursor: catsLoading ? "wait" : "pointer", opacity: (catsLoading || members.length === 0) ? 0.6 : 1 }}>
            {catsLoading ? "🤖 Scanning…" : openCats ? "↺ Re-check Open Categories" : "🔍 Check Open Categories"}
          </button>
          {openCats?.checkedAt && !catsLoading && (
            <span style={{ fontSize: 11, fontWeight: 600, color: daysSinceCatCheck >= 7 ? "#B45309" : "#166534" }}>
              {daysSinceCatCheck >= 7 ? `⏰ Last checked ${daysSinceCatCheck} days ago — time for the weekly refresh` : `✅ Last checked: ${new Date(openCats.checkedAt).toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" })}${daysSinceCatCheck === 0 ? " (today)" : daysSinceCatCheck === 1 ? " (yesterday)" : ` (${daysSinceCatCheck} days ago)`}`}
            </span>
          )}
        </div>
        {catsLoading && (
          <div style={{ display: "flex", alignItems: "center", gap: 8, color: "#6B21A8", fontSize: 12, marginTop: 8 }}>
            <div style={{ width: 14, height: 14, border: "2px solid #7C3AED", borderTopColor: "transparent", borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
            Cross-checking {allCategoriesPresent.length} taken categories against {UAE_OPEN_CATEGORY_POOL.length} categories filled in other UAE chapters…
          </div>
        )}
        {catsError && <div style={{ color: "#991B1B", fontSize: 12, marginTop: 8 }}>⚠️ {catsError}</div>}
        {openCats?.categories?.length > 0 && !catsLoading && (
          <div style={{ border: "1px solid #E9D5FF", borderRadius: 8, background: "#fff", overflow: "hidden" }}>
            {openCats.categories.slice(0, 10).map((c, i) => (
              <div key={i} style={{ display: "flex", gap: 10, padding: "9px 12px", borderBottom: i < openCats.categories.slice(0, 10).length - 1 ? "1px solid #F3F4F6" : "none", background: c.fit === "high" ? "#FFFBEB" : "#fff" }}>
                <div style={{ width: 22, height: 22, borderRadius: "50%", background: c.fit === "high" ? "#8B1A1A" : "#9CA3AF", color: "#fff", fontSize: 11, fontWeight: 800, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>{i + 1}</div>
                <div style={{ minWidth: 0 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                    <span style={{ fontWeight: 800, fontSize: 13, color: "#111" }}>{c.category}</span>
                    {c.fit === "high" && <span style={{ background: "#FEF3C7", color: "#92400E", fontSize: 9, fontWeight: 800, padding: "2px 7px", borderRadius: 10 }}>★ TOP DEMAND</span>}
                    {c.chapters > 0 && <span style={{ background: "#EDE9FE", color: "#5B21B6", fontSize: 9, fontWeight: 700, padding: "2px 7px", borderRadius: 10 }}>📊 {c.chapters} UAE chapter{c.chapters === 1 ? "" : "s"} · {c.memberCount} member{c.memberCount === 1 ? "" : "s"}</span>}
                  </div>
                  {c.synergyWith && <div style={{ fontSize: 11, color: "#4338CA", marginTop: 2 }}>↔ Refers with: {c.synergyWith}</div>}
                  {c.reason && <div style={{ fontSize: 11, color: "#6B7280", fontStyle: "italic", marginTop: 1 }}>{c.reason}</div>}
                </div>
              </div>
            ))}
          </div>
        )}
        {openCats?.categories?.length > 0 && !catsLoading && (
          <div style={{ fontSize: 10.5, color: "#7E22CE", marginTop: 8 }}>🖨️ This list also appears on the print view under the visitor table, ready to share with members.</div>
        )}
      </Card>
    )}

    {viewMode === "list" && showPaste && (
      <Card style={{ marginBottom: 12, background: "#FEFCE8", borderColor: "#F59E0B" }}>
        <div style={{ fontSize: 13, fontWeight: 800, color: "#92400E", marginBottom: 6 }}>📋 Paste visitor list</div>
        <div style={{ fontSize: 11, color: "#B45309", marginBottom: 8, lineHeight: 1.6 }}>
          Copy the visitor rows straight from BNI Connect and paste them below — name, company, specialty, date, phone, email. "Edit" and "Delete" lines are ignored automatically. You'll choose each visitor's inviter in the next step.
        </div>
        <textarea value={pasteText} onChange={e => setPasteText(e.target.value)} rows={10}
          placeholder={"Sarah Tharakan\nNo Company\nCustom Clothing/Tailor\n08/07/2026\n+971 55 470 9460\nsarahbtharakan@gmail.com\nEdit\nDelete\nVik Patel\n..."}
          style={{ width: "100%", boxSizing: "border-box", padding: 10, border: "1px solid #FCD34D", borderRadius: 8, fontSize: 12, fontFamily: "monospace", background: "#fff", resize: "vertical" }} />
        <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
          <button onClick={parsePastedList} disabled={!pasteText.trim()} style={{ background: "#1B2A4A", color: "#fff", border: "none", borderRadius: 8, padding: "8px 18px", fontSize: 12, fontWeight: 700, cursor: pasteText.trim() ? "pointer" : "not-allowed", opacity: pasteText.trim() ? 1 : 0.5 }}>Parse visitors →</button>
          <button onClick={() => { setShowPaste(false); setPasteText(""); }} style={{ background: "#fff", color: "#374151", border: "1px solid #D1D5DB", borderRadius: 8, padding: "8px 16px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>Cancel</button>
        </div>
      </Card>
    )}

    {viewMode === "list" && pastePreview && (
      <Card style={{ marginBottom: 12, background: "#F0FDF4", borderColor: "#22C55E" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
          <div style={{ fontSize: 13, fontWeight: 800, color: "#15803D" }}>✓ {pastePreview.length} visitor{pastePreview.length === 1 ? "" : "s"} detected — assign inviters</div>
          <button onClick={() => setPastePreview(null)} style={{ background: "none", border: "none", fontSize: 14, cursor: "pointer", color: "#15803D" }}>✕</button>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 10, flexWrap: "wrap" }}>
          <span style={{ fontSize: 11, fontWeight: 700, color: "#166534" }}>Same inviter for all:</span>
          <select onChange={e => e.target.value && setAllInviters(e.target.value)} defaultValue=""
            style={{ padding: "6px 10px", border: "1px solid #86EFAC", borderRadius: 8, fontSize: 12, background: "#fff" }}>
            <option value="">Select member...</option>
            {sortedMembers.map(m => <option key={m.id} value={m.name}>{m.name}</option>)}
          </select>
        </div>
        <div style={{ maxHeight: 340, overflowY: "auto", border: "1px solid #BBF7D0", borderRadius: 8, background: "#fff", marginBottom: 10 }}>
          {pastePreview.map((r, i) => {
            const skipped = r.status === "duplicate" || r.status === "error";
            const badge = r.status === "ok" ? { bg: "#D1FAE5", text: "#065F46", label: "✓ Ready" } :
                          r.status === "warning" ? { bg: "#FEF3C7", text: "#92400E", label: "⚠ Check" } :
                          r.status === "duplicate" ? { bg: "#E5E7EB", text: "#374151", label: "⏭ Skip (duplicate)" } :
                          { bg: "#FEE2E2", text: "#991B1B", label: "✗ Skip (error)" };
            return (
              <div key={i} style={{ padding: "10px 12px", borderBottom: "1px solid #F3F4F6", opacity: skipped ? 0.6 : 1 }}>
                <div style={{ display: "flex", justifyContent: "space-between", gap: 8, marginBottom: 6 }}>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontSize: 12, fontWeight: 700 }}>{r.name || <em style={{ color: "#9CA3AF" }}>no name</em>} <span style={{ fontWeight: 400, color: "#6B7280" }}>{r.business || "—"}</span></div>
                    <div style={{ fontSize: 10, color: "#6B7280" }}>{[r.specialty, r.phone, r.email, r.date].filter(Boolean).join(" • ")}</div>
                    {r.warnings.length > 0 && <div style={{ fontSize: 10, color: "#B45309", marginTop: 2 }}>{r.warnings.join(" · ")}</div>}
                  </div>
                  <span style={{ background: badge.bg, color: badge.text, padding: "2px 8px", borderRadius: 10, fontSize: 10, fontWeight: 700, alignSelf: "flex-start", whiteSpace: "nowrap" }}>{badge.label}</span>
                </div>
                {!skipped && (
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                    <div>
                      <label style={{ fontSize: 9, fontWeight: 700, color: "#166534", textTransform: "uppercase", letterSpacing: 0.5 }}>Invited By</label>
                      <select value={r.invitedBy} onChange={e => setPasteField(i, "invitedBy", e.target.value)}
                        style={{ width: "100%", padding: "5px 8px", border: `1px solid ${r.invitedBy ? "#86EFAC" : "#FCD34D"}`, borderRadius: 6, fontSize: 12, background: "#fff" }}>
                        <option value="">Select member...</option>
                        {sortedMembers.map(m => <option key={m.id} value={m.name}>{m.name}</option>)}
                      </select>
                    </div>
                    <div>
                      <label style={{ fontSize: 9, fontWeight: 700, color: "#166534", textTransform: "uppercase", letterSpacing: 0.5 }}>Category (optional)</label>
                      <select value={r.category} onChange={e => setPasteField(i, "category", e.target.value)}
                        style={{ width: "100%", padding: "5px 8px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12, background: "#fff" }}>
                        <option value="">Select category...</option>
                        {pasteCategoryOptions.map(c => <option key={c} value={c}>{c}</option>)}
                      </select>
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <button onClick={confirmPasteImport} disabled={importing || pastePreview.filter(r => r.status === "ok" || r.status === "warning").length === 0}
            style={{ background: "#059669", color: "#fff", border: "none", borderRadius: 8, padding: "8px 18px", fontSize: 12, fontWeight: 700, cursor: importing ? "wait" : "pointer", opacity: importing ? 0.7 : 1 }}>
            {importing ? "Adding…" : `✓ Add ${pastePreview.filter(r => r.status === "ok" || r.status === "warning").length} visitor${pastePreview.filter(r => r.status === "ok" || r.status === "warning").length === 1 ? "" : "s"}`}
          </button>
          <button onClick={() => { setPastePreview(null); setShowPaste(true); }} style={{ background: "#fff", color: "#374151", border: "1px solid #D1D5DB", borderRadius: 8, padding: "8px 16px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>← Back to paste</button>
          {pastePreview.some(r => (r.status === "ok" || r.status === "warning") && !r.invitedBy) &&
            <span style={{ fontSize: 10, color: "#B45309" }}>Some visitors have no inviter selected — they'll be added without one.</span>}
        </div>
      </Card>
    )}

    {viewMode === "list" && importPreview && (
      <Card style={{ marginBottom: 12, background: "#F0F9FF", borderColor: "#38BDF8" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
          <div style={{ fontSize: 13, fontWeight: 800, color: "#0C4A6E" }}>📋 Import Preview — {importPreview.length} row{importPreview.length === 1 ? "" : "s"} found</div>
          <button onClick={() => setImportPreview(null)} style={{ background: "none", border: "none", fontSize: 14, cursor: "pointer", color: "#0C4A6E" }}>✕</button>
        </div>
        {importPreview.length === 0 && <div style={{ fontSize: 12, color: "#6B7280" }}>No data rows found in that file. Fill in the template starting from row 2 and re-upload.</div>}
        {importPreview.length > 0 && <>
          <div style={{ maxHeight: 280, overflowY: "auto", border: "1px solid #BAE6FD", borderRadius: 8, background: "#fff", marginBottom: 10 }}>
            {importPreview.map((r, i) => {
              const badge = r.status === "ok" ? { bg: "#D1FAE5", text: "#065F46", label: "✓ Ready" } :
                            r.status === "warning" ? { bg: "#FEF3C7", text: "#92400E", label: "⚠ Check" } :
                            r.status === "duplicate" ? { bg: "#E5E7EB", text: "#374151", label: "⏭ Skip (duplicate)" } :
                            { bg: "#FEE2E2", text: "#991B1B", label: "✗ Skip (error)" };
              return (
                <div key={i} style={{ display: "flex", justifyContent: "space-between", gap: 8, padding: "8px 10px", borderBottom: "1px solid #F3F4F6", fontSize: 12, opacity: (r.status === "duplicate" || r.status === "error") ? 0.65 : 1 }}>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontWeight: 700 }}>{r.name || <em style={{ color: "#9CA3AF" }}>no name</em>} <span style={{ fontWeight: 400, color: "#6B7280" }}>{r.business}</span></div>
                    <div style={{ fontSize: 10, color: "#6B7280" }}>{[r.phone, r.email, r.invitedBy && `Invited by ${r.invitedBy}`, r.date].filter(Boolean).join(" • ")}</div>
                    {r.warnings.length > 0 && <div style={{ fontSize: 10, color: "#B45309", marginTop: 2 }}>{r.warnings.join(" · ")} (row {r.rowNum})</div>}
                  </div>
                  <span style={{ background: badge.bg, color: badge.text, padding: "2px 8px", borderRadius: 10, fontSize: 10, fontWeight: 700, alignSelf: "flex-start", whiteSpace: "nowrap" }}>{badge.label}</span>
                </div>
              );
            })}
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            <button onClick={confirmImport} disabled={importing || importPreview.filter(r => r.status === "ok" || r.status === "warning").length === 0}
              style={{ background: "#059669", color: "#fff", border: "none", borderRadius: 8, padding: "8px 18px", fontSize: 12, fontWeight: 700, cursor: importing ? "wait" : "pointer", opacity: importing ? 0.7 : 1 }}>
              {importing ? "Importing…" : `✓ Import ${importPreview.filter(r => r.status === "ok" || r.status === "warning").length} visitor${importPreview.filter(r => r.status === "ok" || r.status === "warning").length === 1 ? "" : "s"}`}
            </button>
            <button onClick={() => setImportPreview(null)} style={{ background: "#fff", color: "#374151", border: "1px solid #D1D5DB", borderRadius: 8, padding: "8px 16px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>Cancel</button>
            <span style={{ fontSize: 10, color: "#6B7280" }}>Duplicates and rows with errors are skipped automatically.</span>
          </div>
        </>}
      </Card>
    )}

    {viewMode === "list" && <>
      {showForm && <Card style={{ marginBottom: 12, background: "#FEFCE8" }}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
          {[["name","Full Name"],["business","Business"],["phone","Phone"],["email","Email"]].map(([k,l]) => (
            <div key={k}>
              <label style={{ fontSize: 10, fontWeight: 600 }}>{l}</label>
              <input value={form[k]} onChange={e => setForm(p => ({...p, [k]: e.target.value}))} style={{ width: "100%", padding: "5px 8px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12, boxSizing: "border-box" }} />
            </div>
          ))}
          <div>
            <label style={{ fontSize: 10, fontWeight: 600 }}>Invited By</label>
            <select value={form.invitedBy} onChange={e => setForm(p => ({...p, invitedBy: e.target.value}))} style={{ width: "100%", padding: "5px 8px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12 }}>
              <option value="">Select member...</option>
              {members.map(m => <option key={m.id} value={m.name}>{m.name}</option>)}
            </select>
          </div>
          <div>
            <label style={{ fontSize: 10, fontWeight: 600 }}>Category</label>
            <select value={form.category} onChange={e => setForm(p => ({...p, category: e.target.value}))} style={{ width: "100%", padding: "5px 8px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12 }}>
              <option value="">Select category...</option>
              {allCategoriesPresent.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <label style={{ fontSize: 10, fontWeight: 600 }}>Specialty</label>
            <input value={form.specialty} onChange={e => setForm(p => ({...p, specialty: e.target.value}))} style={{ width: "100%", padding: "5px 8px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12, boxSizing: "border-box" }} />
          </div>
          <div>
            <label style={{ fontSize: 10, fontWeight: 600 }}>Meeting Date</label>
            <input type="date" value={form.date} onChange={e => setForm(p => ({...p, date: e.target.value}))} style={{ width: "100%", padding: "5px 8px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12, boxSizing: "border-box" }} />
          </div>
        </div>
        <button onClick={addVisitor} style={{ marginTop: 10, background: "#059669", color: "#fff", border: "none", borderRadius: 8, padding: "7px 16px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>Save</button>
      </Card>}

      {visitors.map(v => {
        const isExpanded = expandedId === v.id;
        const isEditing = editingId === v.id;
        const topMatch = findMatches(v, asks, members)[0];
        const verdict = v.validation?.verdict;
        const verdictBadge = verdict === "GREEN" ? { bg: "#D1FAE5", text: "#065F46", label: "✅ Green" } :
                             verdict === "AMBER" ? { bg: "#FEF3C7", text: "#92400E", label: "⚠️ Amber" } :
                             verdict === "RED"   ? { bg: "#FEE2E2", text: "#991B1B", label: "🛑 Red" } : null;
        const panel = activePanel[v.id]; // 'brief' or 'validate'

        if (isEditing) {
          return (
            <Card key={v.id} style={{ marginBottom: 10, padding: 14, background: "#FEFCE8", border: "2px solid #F59E0B" }}>
              <div style={{ fontSize: 12, fontWeight: 800, color: "#92400E", marginBottom: 8 }}>✏️ Editing visitor</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                {[["name","Full Name"],["business","Business"],["phone","Phone"],["email","Email"]].map(([k,l]) => (
                  <div key={k}>
                    <label style={{ fontSize: 10, fontWeight: 600 }}>{l}</label>
                    <input value={editForm[k] || ""} onChange={e => setEditForm(p => ({...p, [k]: e.target.value}))} style={{ width: "100%", padding: "5px 8px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12, boxSizing: "border-box" }} />
                  </div>
                ))}
                <div>
                  <label style={{ fontSize: 10, fontWeight: 600 }}>Invited By</label>
                  <select value={editForm.invitedBy || ""} onChange={e => setEditForm(p => ({...p, invitedBy: e.target.value}))} style={{ width: "100%", padding: "5px 8px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12 }}>
                    <option value="">Select member...</option>
                    {members.map(m => <option key={m.id} value={m.name}>{m.name}</option>)}
                  </select>
                </div>
                <div>
                  <label style={{ fontSize: 10, fontWeight: 600 }}>Category</label>
                  <select value={editForm.category || ""} onChange={e => setEditForm(p => ({...p, category: e.target.value}))} style={{ width: "100%", padding: "5px 8px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12 }}>
                    <option value="">Select category...</option>
                    {allCategoriesPresent.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
                <div>
                  <label style={{ fontSize: 10, fontWeight: 600 }}>Specialty</label>
                  <input value={editForm.specialty || ""} onChange={e => setEditForm(p => ({...p, specialty: e.target.value}))} style={{ width: "100%", padding: "5px 8px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12, boxSizing: "border-box" }} />
                </div>
                <div>
                  <label style={{ fontSize: 10, fontWeight: 600 }}>Meeting Date</label>
                  <input type="date" value={editForm.date || ""} onChange={e => setEditForm(p => ({...p, date: e.target.value}))} style={{ width: "100%", padding: "5px 8px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12, boxSizing: "border-box" }} />
                </div>
                <div style={{ gridColumn: "span 2" }}>
                  <label style={{ fontSize: 10, fontWeight: 600 }}>Call notes</label>
                  <textarea value={editForm.callNotes || ""} onChange={e => setEditForm(p => ({...p, callNotes: e.target.value}))} rows={2} style={{ width: "100%", padding: "5px 8px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12, boxSizing: "border-box", resize: "vertical", fontFamily: "inherit" }} />
                </div>
              </div>
              <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                <button onClick={saveEdit} style={{ background: "#059669", color: "#fff", border: "none", borderRadius: 8, padding: "7px 16px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>💾 Save Changes</button>
                <button onClick={cancelEdit} style={{ background: "#fff", color: "#374151", border: "1px solid #D1D5DB", borderRadius: 8, padding: "7px 16px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>Cancel</button>
              </div>
            </Card>
          );
        }

        return (
          <Card key={v.id} style={{ marginBottom: 8, padding: 12, border: v.bio ? "1px solid #C7D2FE" : "1px solid #E5E7EB" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                  <div style={{ fontWeight: 700, fontSize: 13 }}>{v.name}</div>
                  {v.bio && <span style={{ background: "#EEF2FF", color: "#4338CA", fontSize: 9, fontWeight: 700, padding: "2px 7px", borderRadius: 20 }}>✦ Briefed</span>}
                  {verdictBadge && <span style={{ background: verdictBadge.bg, color: verdictBadge.text, fontSize: 9, fontWeight: 700, padding: "2px 7px", borderRadius: 20 }}>{verdictBadge.label}</span>}
                  {topMatch && <span style={{ background: topMatch.score >= 70 ? "#FEF3C7" : "#DBEAFE", color: topMatch.score >= 70 ? "#92400E" : "#1E40AF", fontSize: 9, fontWeight: 700, padding: "2px 7px", borderRadius: 20 }}>
                    → {topMatch.member?.name.split(" ")[0]}
                  </span>}
                </div>
                <div style={{ fontSize: 11, color: "#6B7280" }}>{v.business} • {v.category || "—"}</div>
                <div style={{ fontSize: 10, color: "#9CA3AF" }}>Invited by {v.invitedBy || "—"} • {v.date}</div>
              </div>
              <div style={{ display: "flex", gap: 4, alignItems: "center", flexShrink: 0, flexWrap: "wrap", justifyContent: "flex-end" }}>
                <StatusBadge status={v.status} />
                <select value={v.status} onChange={e => updateStatus(v.id, e.target.value)} style={{ fontSize: 10, padding: "2px 4px", border: "1px solid #D1D5DB", borderRadius: 4 }}>
                  {Object.keys(STATUS_COLORS).map(s => <option key={s} value={s}>{STATUS_COLORS[s].label}</option>)}
                </select>
              </div>
            </div>

            {/* Action row */}
            <div style={{ display: "flex", gap: 6, marginTop: 10, flexWrap: "wrap" }}>
              <button
                onClick={() => { setExpandedId(isExpanded && panel === "brief" ? null : v.id); setActivePanel(p => ({ ...p, [v.id]: "brief" })); }}
                style={{ background: panel === "brief" && isExpanded ? "#EEF2FF" : "#F9FAFB", border: "1px solid #C7D2FE", color: "#4338CA", borderRadius: 6, padding: "5px 10px", fontSize: 11, cursor: "pointer", fontWeight: 700 }}>
                ✨ Brief
              </button>
              <button
                onClick={() => { setExpandedId(isExpanded && panel === "validate" ? null : v.id); setActivePanel(p => ({ ...p, [v.id]: "validate" })); }}
                style={{ background: panel === "validate" && isExpanded ? "#D1FAE5" : "#F9FAFB", border: "1px solid #6EE7B7", color: "#065F46", borderRadius: 6, padding: "5px 10px", fontSize: 11, cursor: "pointer", fontWeight: 700 }}>
                🛡️ Validate
              </button>
              <button
                onClick={() => startEdit(v)}
                style={{ background: "#F9FAFB", border: "1px solid #FCD34D", color: "#92400E", borderRadius: 6, padding: "5px 10px", fontSize: 11, cursor: "pointer", fontWeight: 700 }}>
                ✏️ Edit
              </button>
              <button
                onClick={() => requestArchive(v)}
                style={{ background: "#F9FAFB", border: "1px solid #C4B5FD", color: "#5B21B6", borderRadius: 6, padding: "5px 10px", fontSize: 11, cursor: "pointer", fontWeight: 700 }}>
                🗄️ Archive
              </button>
              <button
                onClick={() => requestDelete(v)}
                style={{ background: "#fff", border: "1px solid #FCA5A5", color: "#991B1B", borderRadius: 6, padding: "5px 10px", fontSize: 11, cursor: "pointer", fontWeight: 700 }}>
                🗑️ Delete
              </button>
            </div>

            {isExpanded && panel === "brief" && (
              <div style={{ marginTop: 10, borderTop: "1px solid #F3F4F6", paddingTop: 10 }}>
                <VisitorIntelligence visitor={v} onBioSaved={saveBio} />
              </div>
            )}
            {isExpanded && panel === "validate" && (
              <div style={{ marginTop: 10, borderTop: "1px solid #F3F4F6", paddingTop: 10 }}>
                <VisitorValidation visitor={v} members={members} onValidationSaved={saveValidation} />
              </div>
            )}
          </Card>
        );
      })}
    </>}

    {viewMode === "print" && <>
      <Card style={{ marginBottom: 12, background: "#F0FDF4", borderColor: "#22C55E", padding: 14 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
          <div>
            <label style={{ fontSize: 11, fontWeight: 700, color: "#15803D", display: "block", marginBottom: 4 }}>📅 Meeting Date</label>
            <select value={printDate} onChange={e => setPrintDate(e.target.value)} style={{ padding: "6px 10px", border: "1px solid #BBF7D0", borderRadius: 8, fontSize: 12, background: "#fff", color: "#111", minWidth: 160 }}>
              {allDates.map(d => <option key={d} value={d}>{d}</option>)}
            </select>
            <input type="date" value={printDate} onChange={e => setPrintDate(e.target.value)} style={{ marginLeft: 8, padding: "6px 8px", border: "1px solid #BBF7D0", borderRadius: 8, fontSize: 12, background: "#fff" }} />
          </div>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 11, color: "#166534" }}><strong>{printVisitors.length}</strong> visitor{printVisitors.length !== 1 ? "s" : ""} on this date</div>
          </div>
          <button
            onClick={runPrintAIMatch}
            disabled={matchLoading || printVisitors.length === 0}
            title="AI matches each visitor to members using open asks and the full member list"
            style={{ background: "#7C3AED", color: "#fff", border: "none", borderRadius: 8, padding: "9px 20px", fontSize: 13, fontWeight: 700, cursor: (matchLoading || printVisitors.length === 0) ? "wait" : "pointer", opacity: (matchLoading || printVisitors.length === 0) ? 0.6 : 1, minWidth: 160 }}>
            {matchLoading ? "🤖 Matching…" : Object.keys(printMatches).length > 0 ? "↺ Re-run AI Match" : "🤖 AI-Match Visitors"}
          </button>
          <button onClick={() => setShowPrintModal(true)} style={{ background: "#8B1A1A", color: "#fff", border: "none", borderRadius: 8, padding: "9px 20px", fontSize: 13, fontWeight: 700, cursor: "pointer", minWidth: 160 }}>
            🖨️ Open Print View
          </button>
        </div>
        {matchLoading && (
          <div style={{ display: "flex", alignItems: "center", gap: 8, color: "#5B21B6", fontSize: 12, marginTop: 10 }}>
            <div style={{ width: 14, height: 14, border: "2px solid #7C3AED", borderTopColor: "transparent", borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
            Matching {printVisitors.length} visitor{printVisitors.length !== 1 ? "s" : ""} against {asks.filter(a => a.status === "open").length} open asks and {members.length} members…
          </div>
        )}
        {matchError && <div style={{ color: "#991B1B", fontSize: 12, marginTop: 8 }}>⚠️ {matchError}</div>}
        {!matchLoading && Object.keys(printMatches).length > 0 && (
          <div style={{ fontSize: 11, color: "#166534", marginTop: 8 }}>✅ AI matches loaded — they now appear in the "BNI Matches — Introduce To" column below and in the print view. (Session only — re-run after a page refresh.)</div>
        )}
      </Card>
      <div style={{ border: "2px solid #E5E7EB", borderRadius: 10, overflow: "hidden", background: "#fff" }}>
        <div style={{ background: "#374151", color: "#9CA3AF", fontSize: 10, padding: "5px 14px", fontWeight: 600, letterSpacing: 1, textTransform: "uppercase" }}>Preview — A4 Landscape</div>
        <div id="bni-print-area">
          <PrintableVisitorList visitors={printVisitors} meetingDate={printDate} asks={asks} members={members} aiMatches={printMatches} openCategories={openCats} />
        </div>
      </div>
    </>}

    {showPrintModal && (
      <div style={{ position: "fixed", inset: 0, zIndex: 9999, background: "#fff", overflowY: "auto" }}>
        <div className="no-print" style={{ position: "sticky", top: 0, zIndex: 10000, background: "#1B2A4A", color: "#fff", padding: "10px 20px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap", boxShadow: "0 2px 12px rgba(0,0,0,0.3)" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <span style={{ fontSize: 15, fontWeight: 800 }}>🖨️ Print-Ready View</span>
            <span style={{ fontSize: 12, color: "#93C5FD" }}>
              Press <kbd style={{ background: "#374151", border: "1px solid #4B5563", borderRadius: 4, padding: "1px 6px", fontFamily: "monospace", fontSize: 12 }}>Ctrl+P</kbd>
              &nbsp;or <kbd style={{ background: "#374151", border: "1px solid #4B5563", borderRadius: 4, padding: "1px 6px", fontFamily: "monospace", fontSize: 12 }}>⌘+P</kbd>
              &nbsp;→ <strong style={{ color: "#FCD34D" }}>"Save as PDF"</strong> → <strong style={{ color: "#FCD34D" }}>Layout: Landscape</strong>
            </span>
          </div>
          <button onClick={() => setShowPrintModal(false)} style={{ background: "#8B1A1A", color: "#fff", border: "none", borderRadius: 8, padding: "7px 16px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>✕ Close</button>
        </div>
        <div id="bni-modal-print-area">
          <PrintableVisitorList visitors={printVisitors} meetingDate={printDate} asks={asks} members={members} aiMatches={printMatches} openCategories={openCats} />
        </div>
        <style>{`
          @media print {
            .no-print { display: none !important; }
            body { margin: 0 !important; }
            @page { size: A4 landscape; margin: 10mm 12mm; }
          }
        `}</style>
      </div>
    )}
  </div>;
}

// ═══════════════════════════════════════════
// NEW: ARCHIVE TAB — month-based lookup
// ═══════════════════════════════════════════
function ArchiveTab({ archived, setArchived, visitors, setVisitors }) {
  const [filterMonth, setFilterMonth] = useState("all");
  const [search, setSearch] = useState("");
  const [confirmDelete, setConfirmDelete] = useState(null);
  const [confirmRestore, setConfirmRestore] = useState(null);

  // Build month index from the visitor's meeting date (more useful than archivedAt)
  const monthsInArchive = [...new Set(archived.map(v => (v.date || "").slice(0, 7)).filter(Boolean))].sort().reverse();

  const monthLabel = (ym) => {
    if (!ym) return "Unknown";
    const [y, m] = ym.split("-");
    return new Date(parseInt(y), parseInt(m) - 1, 1).toLocaleDateString("en-GB", { month: "long", year: "numeric" });
  };

  const filtered = archived.filter(v => {
    const monthOk = filterMonth === "all" || (v.date || "").startsWith(filterMonth);
    const searchOk = !search || v.name.toLowerCase().includes(search.toLowerCase()) ||
                     v.business.toLowerCase().includes(search.toLowerCase()) ||
                     (v.invitedBy || "").toLowerCase().includes(search.toLowerCase());
    return monthOk && searchOk;
  }).sort((a, b) => (b.date || "").localeCompare(a.date || ""));

  // Group by month for display
  const grouped = {};
  filtered.forEach(v => {
    const month = (v.date || "").slice(0, 7) || "unknown";
    if (!grouped[month]) grouped[month] = [];
    grouped[month].push(v);
  });

  const archiveAllPast = () => {
    const today = new Date().toISOString().split("T")[0];
    const toArchive = visitors.filter(v => v.date < today);
    if (toArchive.length === 0) {
      alert("No visitors with past meeting dates to archive.");
      return;
    }
    if (!confirm(`Archive ${toArchive.length} visitor(s) with past meeting dates?`)) return;
    const archivedAt = new Date().toISOString();
    setArchived(p => [...p, ...toArchive.map(v => ({ ...v, archivedAt }))]);
    setVisitors(p => p.filter(v => v.date >= today));
  };

  const doRestore = () => {
    setVisitors(p => [...p, { ...confirmRestore }]);
    setArchived(p => p.filter(v => v.id !== confirmRestore.id));
    setConfirmRestore(null);
  };

  const doDelete = () => {
    setArchived(p => p.filter(v => v.id !== confirmDelete.id));
    setConfirmDelete(null);
  };

  return (
    <div>
      <ConfirmModal
        open={!!confirmDelete}
        title="Permanently delete from archive?"
        message={confirmDelete ? `"${confirmDelete.name}" will be permanently deleted from your archive. This cannot be undone.` : ""}
        confirmLabel="Yes, delete forever"
        onConfirm={doDelete}
        onCancel={() => setConfirmDelete(null)}
      />
      <ConfirmModal
        open={!!confirmRestore}
        title="Restore visitor to active list?"
        message={confirmRestore ? `"${confirmRestore.name}" will move from the archive back to your active visitors list.` : ""}
        confirmLabel="Yes, restore"
        danger={false}
        onConfirm={doRestore}
        onCancel={() => setConfirmRestore(null)}
      />

      <Card style={{ marginBottom: 12, background: "linear-gradient(135deg, #F5F3FF 0%, #EDE9FE 100%)", borderColor: "#A78BFA" }}>
        <div style={{ fontSize: 12, fontWeight: 700, color: "#5B21B6", marginBottom: 4 }}>🗄️ Visitor Archive</div>
        <div style={{ fontSize: 11, color: "#6D28D9", lineHeight: 1.6 }}>
          Past visitors stored by meeting month. Use this to look up who visited, find returning prospects, or pull historical records. Archived visitors don't appear in dashboard counts, AI matching, or the seat planner.
        </div>
      </Card>

      <Card style={{ marginBottom: 12 }}>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "flex-end" }}>
          <div style={{ flex: 1, minWidth: 180 }}>
            <label style={{ fontSize: 10, fontWeight: 700, color: "#374151", display: "block", marginBottom: 3 }}>📅 Filter by month</label>
            <select value={filterMonth} onChange={e => setFilterMonth(e.target.value)} style={{ width: "100%", padding: "6px 10px", border: "1px solid #D1D5DB", borderRadius: 8, fontSize: 12 }}>
              <option value="all">All months ({archived.length} total)</option>
              {monthsInArchive.map(m => (
                <option key={m} value={m}>{monthLabel(m)} ({archived.filter(v => (v.date || "").startsWith(m)).length})</option>
              ))}
            </select>
          </div>
          <div style={{ flex: 1, minWidth: 180 }}>
            <label style={{ fontSize: 10, fontWeight: 700, color: "#374151", display: "block", marginBottom: 3 }}>🔍 Search</label>
            <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Name, business, or inviter..." style={{ width: "100%", padding: "6px 10px", border: "1px solid #D1D5DB", borderRadius: 8, fontSize: 12, boxSizing: "border-box" }} />
          </div>
          <button onClick={archiveAllPast} style={{ background: "#5B21B6", color: "#fff", border: "none", borderRadius: 8, padding: "7px 14px", fontSize: 11, fontWeight: 700, cursor: "pointer", whiteSpace: "nowrap" }}>
            ⚡ Archive all past meetings
          </button>
        </div>
      </Card>

      {filtered.length === 0 ? (
        <Card style={{ textAlign: "center", padding: 32, color: "#9CA3AF" }}>
          <div style={{ fontSize: 32, marginBottom: 8 }}>🗂️</div>
          <div style={{ fontSize: 13, fontWeight: 600, color: "#6B7280" }}>
            {archived.length === 0 ? "Archive is empty" : "No visitors match your filter"}
          </div>
          <div style={{ fontSize: 11, marginTop: 6 }}>
            {archived.length === 0 ? "Archive visitors from the Visitors tab to keep historical records." : "Try a different month or clear your search."}
          </div>
        </Card>
      ) : (
        Object.keys(grouped).sort().reverse().map(month => (
          <div key={month} style={{ marginBottom: 14 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
              <div style={{ background: "#5B21B6", color: "#fff", padding: "4px 12px", borderRadius: 16, fontSize: 11, fontWeight: 800 }}>
                {monthLabel(month)}
              </div>
              <div style={{ fontSize: 10, color: "#9CA3AF" }}>{grouped[month].length} visitor{grouped[month].length !== 1 ? "s" : ""}</div>
              <div style={{ flex: 1, height: 1, background: "#E5E7EB" }} />
            </div>
            {grouped[month].map(v => (
              <Card key={v.id} style={{ marginBottom: 6, padding: 10, borderLeft: "3px solid #A78BFA" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
                      <span style={{ fontWeight: 700, fontSize: 12 }}>{v.name}</span>
                      <StatusBadge status={v.status} />
                      {v.validation?.verdict && (
                        <span style={{ fontSize: 9, fontWeight: 700, padding: "1px 6px", borderRadius: 8,
                          background: v.validation.verdict === "GREEN" ? "#D1FAE5" : v.validation.verdict === "AMBER" ? "#FEF3C7" : "#FEE2E2",
                          color: v.validation.verdict === "GREEN" ? "#065F46" : v.validation.verdict === "AMBER" ? "#92400E" : "#991B1B" }}>
                          {v.validation.verdict}
                        </span>
                      )}
                    </div>
                    <div style={{ fontSize: 11, color: "#6B7280" }}>{v.business} • {v.category || "—"}</div>
                    <div style={{ fontSize: 10, color: "#9CA3AF" }}>
                      Meeting: {v.date} • Invited by {v.invitedBy || "walk-in"} • {v.phone || "no phone"}
                    </div>
                  </div>
                  <div style={{ display: "flex", gap: 4, flexShrink: 0 }}>
                    <button onClick={() => setConfirmRestore(v)} style={{ background: "#EEF2FF", border: "1px solid #C7D2FE", color: "#4338CA", borderRadius: 6, padding: "4px 8px", fontSize: 10, cursor: "pointer", fontWeight: 700 }}>↩ Restore</button>
                    <button onClick={() => setConfirmDelete(v)} style={{ background: "#fff", border: "1px solid #FCA5A5", color: "#991B1B", borderRadius: 6, padding: "4px 8px", fontSize: 10, cursor: "pointer", fontWeight: 700 }}>🗑️</button>
                  </div>
                </div>
              </Card>
            ))}
          </div>
        ))
      )}
    </div>
  );
}

function AsksTab({ asks, setAsks, members }) {
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({ memberId: "", askType: "general_role", targetName: "", targetCompany: "", targetCategory: "", targetRole: "", notes: "" });
  const [confirmDelete, setConfirmDelete] = useState(null);

  const allCategories = [...new Set(members.map(m => m.category))].sort();

  const addAsk = async () => {
    const member = members.find(m => m.id === Number(form.memberId));
    if (!member) return;
    const { data } = await supabase.from("asks").insert([{
      member_id: member.id, member_name: member.name, ask_type: form.askType,
      target_name: form.targetName, target_company: form.targetCompany,
      target_category: form.targetCategory, target_role: form.targetRole,
      notes: form.notes, date: new Date().toISOString().split("T")[0], status: "open"
    }]).select();
    if (data?.[0]) setAsks(p => [...p, { ...data[0], memberId: data[0].member_id, memberName: data[0].member_name, askType: data[0].ask_type, targetName: data[0].target_name, targetCompany: data[0].target_company, targetCategory: data[0].target_category, targetRole: data[0].target_role }]);
    setForm({ memberId: "", askType: "general_role", targetName: "", targetCompany: "", targetCategory: "", targetRole: "", notes: "" });
    setShowForm(false);
  };
  const closeAsk = async (id) => {
    await supabase.from("asks").update({ status: "fulfilled" }).eq("id", id);
    setAsks(p => p.map(a => a.id === id ? { ...a, status: "fulfilled" } : a));
  };
  const doDelete = async () => { await supabase.from("asks").delete().eq("id", confirmDelete.id); setAsks(p => p.filter(a => a.id !== confirmDelete.id)); setConfirmDelete(null); };

  // ═══════════════════════════════════════════
  // PASTE ASKS — paste one or many members' ask lists; member headers auto-detected
  // ═══════════════════════════════════════════
  const [showPasteAsks, setShowPasteAsks] = useState(false);
  const [askPasteText, setAskPasteText] = useState("");
  const [askPreview, setAskPreview] = useState(null); // [{memberId, asks:[text]}]
  const [savingAsks, setSavingAsks] = useState(false);
  const sortedMembers = [...members].sort((a, b) => a.name.localeCompare(b.name));

  const parseAsksPaste = () => {
    const bulletRe = /^\s*([*\-•·▪‣►>]|\d+[.)])\s+/;
    const norm = (s) => s.toLowerCase().replace(/[^a-z\s]/g, "").replace(/\s+/g, " ").trim();
    const findMember = (line) => {
      const cleaned = line.replace(/['’]s\b/gi, "").replace(/\basks?\b/gi, "").replace(/[–—\-:|]+/g, " ");
      const n = norm(cleaned);
      if (!n) return null;
      let m = members.find(mm => norm(mm.name) === n);
      if (m) return m;
      m = members.find(mm => n.startsWith(norm(mm.name)));
      if (m) return m;
      const words = n.split(" ").filter(w => w.length > 1);
      if (words.length >= 2) {
        m = members.find(mm => { const mn = norm(mm.name); return words.every(w => mn.includes(w)); });
        if (m) return m;
      }
      return null;
    };

    const lines = askPasteText.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    const groups = [];
    let cur = null;
    for (const raw of lines) {
      const isBullet = bulletRe.test(raw);
      if (!isBullet) {
        const m = findMember(raw);
        if (m) { cur = { memberId: String(m.id), asks: [] }; groups.push(cur); continue; }
      }
      const text = raw.replace(bulletRe, "").trim();
      if (!text) continue;
      if (!cur) { cur = { memberId: "", asks: [] }; groups.push(cur); }
      cur.asks.push(text);
    }
    const cleaned = groups.filter(g => g.asks.length > 0);
    if (!cleaned.length) { alert("Couldn't detect any asks in that text. Paste the list with one ask per line."); return; }
    setAskPreview(cleaned);
    setShowPasteAsks(false);
  };

  const setGroupMember = (gi, memberId) => setAskPreview(p => p.map((g, i) => i === gi ? { ...g, memberId } : g));
  const removeAskFromGroup = (gi, ai) => setAskPreview(p =>
    p.map((g, i) => i === gi ? { ...g, asks: g.asks.filter((_, j) => j !== ai) } : g).filter(g => g.asks.length > 0)
  );

  const totalPreviewAsks = askPreview ? askPreview.reduce((s, g) => s + g.asks.length, 0) : 0;
  const unassignedGroups = askPreview ? askPreview.filter(g => !g.memberId).length : 0;

  const confirmAskImport = async () => {
    setSavingAsks(true);
    try {
      const today = toYMD(new Date());
      const rows = askPreview.flatMap(g => {
        const member = members.find(m => String(m.id) === String(g.memberId));
        if (!member) return [];
        return g.asks.map(text => ({
          member_id: member.id, member_name: member.name, ask_type: "free_text",
          target_name: "", target_company: "", target_category: "", target_role: "",
          notes: text, date: today, status: "open",
        }));
      });
      const { data, error } = await supabase.from("asks").insert(rows).select();
      if (error) throw error;
      if (data?.length) {
        setAsks(p => [...p, ...data.map(a => ({
          ...a, memberId: a.member_id, memberName: a.member_name, askType: a.ask_type,
          targetName: a.target_name, targetCompany: a.target_company,
          targetCategory: a.target_category, targetRole: a.target_role,
        }))]);
      }
      setAskPreview(null);
      setAskPasteText("");
    } catch (err) {
      console.error("Ask import failed:", err);
      alert("Import failed — nothing was saved. Check your connection and try again.");
    }
    setSavingAsks(false);
  };

  return <div>
    <ConfirmModal open={!!confirmDelete} title="Delete this ask?" message={confirmDelete ? `Delete the ask from ${confirmDelete.memberName}? This cannot be undone.` : ""} confirmLabel="Yes, delete" onConfirm={doDelete} onCancel={() => setConfirmDelete(null)} />

    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
      <span style={{ fontWeight: 700, fontSize: 15 }}>Member Asks Database ({asks.length})</span>
      <div style={{ display: "flex", gap: 6 }}>
        <button onClick={() => { setShowPasteAsks(!showPasteAsks); setAskPreview(null); }} style={{ background: showPasteAsks ? "#1B2A4A" : "#fff", color: showPasteAsks ? "#fff" : "#1B2A4A", border: "1px solid #1B2A4A", borderRadius: 8, padding: "7px 12px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>📋 Paste Asks</button>
        <button onClick={() => setShowForm(!showForm)} style={{ background: "#8B1A1A", color: "#fff", border: "none", borderRadius: 8, padding: "7px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>+ New Ask</button>
      </div>
    </div>

    {showPasteAsks && (
      <Card style={{ marginBottom: 12, background: "#FEFCE8", borderColor: "#F59E0B" }}>
        <div style={{ fontSize: 13, fontWeight: 800, color: "#92400E", marginBottom: 6 }}>📋 Paste member asks</div>
        <div style={{ fontSize: 11, color: "#B45309", marginBottom: 8, lineHeight: 1.6 }}>
          One ask per line — companies, people, roles, anything. Start a block with the member's name (e.g. "Jesika Menon – Asks") and every line under it is assigned to them. You can paste several members' lists in one go; you'll confirm each member in the next step.
        </div>
        <textarea value={askPasteText} onChange={e => setAskPasteText(e.target.value)} rows={10}
          placeholder={"Jesika Menon – Asks\n* Nakheel\n* Emaar – Community Manager\n* Repton School\n\nMohit Sharma – Asks\n* Business owners with AED 500K+ savings\n* DIFC family offices"}
          style={{ width: "100%", boxSizing: "border-box", padding: 10, border: "1px solid #FCD34D", borderRadius: 8, fontSize: 12, fontFamily: "monospace", background: "#fff", resize: "vertical" }} />
        <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
          <button onClick={parseAsksPaste} disabled={!askPasteText.trim()} style={{ background: "#1B2A4A", color: "#fff", border: "none", borderRadius: 8, padding: "8px 18px", fontSize: 12, fontWeight: 700, cursor: askPasteText.trim() ? "pointer" : "not-allowed", opacity: askPasteText.trim() ? 1 : 0.5 }}>Parse asks →</button>
          <button onClick={() => { setShowPasteAsks(false); setAskPasteText(""); }} style={{ background: "#fff", color: "#374151", border: "1px solid #D1D5DB", borderRadius: 8, padding: "8px 16px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>Cancel</button>
        </div>
      </Card>
    )}

    {askPreview && (
      <Card style={{ marginBottom: 12, background: "#F0FDF4", borderColor: "#22C55E" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
          <div style={{ fontSize: 13, fontWeight: 800, color: "#15803D" }}>✓ {totalPreviewAsks} ask{totalPreviewAsks === 1 ? "" : "s"} detected across {askPreview.length} member{askPreview.length === 1 ? "" : "s"}</div>
          <button onClick={() => setAskPreview(null)} style={{ background: "none", border: "none", fontSize: 14, cursor: "pointer", color: "#15803D" }}>✕</button>
        </div>
        {askPreview.map((g, gi) => (
          <div key={gi} style={{ border: "1px solid #BBF7D0", borderRadius: 8, background: "#fff", padding: 10, marginBottom: 8 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8, flexWrap: "wrap" }}>
              <span style={{ fontSize: 10, fontWeight: 700, color: "#166534", textTransform: "uppercase", letterSpacing: 0.5 }}>Member:</span>
              <select value={g.memberId} onChange={e => setGroupMember(gi, e.target.value)}
                style={{ padding: "5px 10px", border: `1px solid ${g.memberId ? "#86EFAC" : "#FCA5A5"}`, borderRadius: 6, fontSize: 12, background: "#fff", fontWeight: 600 }}>
                <option value="">⚠ Select member...</option>
                {sortedMembers.map(m => <option key={m.id} value={String(m.id)}>{m.name} ({m.specialty})</option>)}
              </select>
              <span style={{ fontSize: 10, color: "#6B7280" }}>{g.asks.length} ask{g.asks.length === 1 ? "" : "s"}</span>
            </div>
            <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
              {g.asks.map((t, ai) => (
                <span key={ai} style={{ display: "inline-flex", alignItems: "center", gap: 6, background: "#F5F3FF", border: "1px solid #DDD6FE", borderRadius: 20, padding: "4px 10px", fontSize: 11, color: "#5B21B6" }}>
                  🎯 {t}
                  <button onClick={() => removeAskFromGroup(gi, ai)} title="Remove this ask" style={{ background: "none", border: "none", cursor: "pointer", color: "#8B5CF6", fontSize: 12, padding: 0, lineHeight: 1 }}>✕</button>
                </span>
              ))}
            </div>
          </div>
        ))}
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <button onClick={confirmAskImport} disabled={savingAsks || unassignedGroups > 0 || totalPreviewAsks === 0}
            style={{ background: "#059669", color: "#fff", border: "none", borderRadius: 8, padding: "8px 18px", fontSize: 12, fontWeight: 700, cursor: savingAsks ? "wait" : (unassignedGroups > 0 ? "not-allowed" : "pointer"), opacity: (savingAsks || unassignedGroups > 0) ? 0.6 : 1 }}>
            {savingAsks ? "Saving…" : `✓ Save ${totalPreviewAsks} ask${totalPreviewAsks === 1 ? "" : "s"}`}
          </button>
          <button onClick={() => { setAskPreview(null); setShowPasteAsks(true); }} style={{ background: "#fff", color: "#374151", border: "1px solid #D1D5DB", borderRadius: 8, padding: "8px 16px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>← Back to paste</button>
          {unassignedGroups > 0 && <span style={{ fontSize: 10, color: "#991B1B" }}>Select a member for every group before saving.</span>}
        </div>
      </Card>
    )}

    <Card style={{ marginBottom: 12, background: "#EEF2FF", borderColor: "#818CF8" }}>
      <div style={{ fontSize: 12, fontWeight: 700, color: "#4338CA", marginBottom: 4 }}>🎯 How Asks Work</div>
      <div style={{ fontSize: 11, color: "#3730A3", lineHeight: 1.6 }}>
        Record member asks each week from their 60-second presentations. Asks can be:<br/>
        <strong>Specific Person</strong> — e.g. "Pallavi Dean from Roar" | <strong>Specific Company</strong> — e.g. "Pixl Global" | <strong>General Role</strong> — e.g. "CFOs of companies"
      </div>
    </Card>

    {showForm && <Card style={{ marginBottom: 12, background: "#FEFCE8" }}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
        <div>
          <label style={{ fontSize: 10, fontWeight: 600 }}>Member</label>
          <select value={form.memberId} onChange={e => setForm(p => ({...p, memberId: e.target.value}))} style={{ width: "100%", padding: "5px 8px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12 }}>
            <option value="">Select member...</option>
            {members.map(m => <option key={m.id} value={m.id}>{m.name} ({m.specialty})</option>)}
          </select>
        </div>
        <div>
          <label style={{ fontSize: 10, fontWeight: 600 }}>Ask Type</label>
          <select value={form.askType} onChange={e => setForm(p => ({...p, askType: e.target.value}))} style={{ width: "100%", padding: "5px 8px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12 }}>
            <option value="specific_person">Specific Person</option>
            <option value="specific_company">Specific Company</option>
            <option value="general_role">General Role / Type</option>
          </select>
        </div>
        {form.askType === "specific_person" && <>
          <div><label style={{ fontSize: 10, fontWeight: 600 }}>Person Name</label><input value={form.targetName} onChange={e => setForm(p => ({...p, targetName: e.target.value}))} style={{ width: "100%", padding: "5px 8px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12, boxSizing: "border-box" }} placeholder="e.g. Pallavi Dean" /></div>
          <div><label style={{ fontSize: 10, fontWeight: 600 }}>Company</label><input value={form.targetCompany} onChange={e => setForm(p => ({...p, targetCompany: e.target.value}))} style={{ width: "100%", padding: "5px 8px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12, boxSizing: "border-box" }} placeholder="e.g. Roar" /></div>
        </>}
        {form.askType === "specific_company" && <div style={{ gridColumn: "span 2" }}>
          <label style={{ fontSize: 10, fontWeight: 600 }}>Company Name</label>
          <input value={form.targetCompany} onChange={e => setForm(p => ({...p, targetCompany: e.target.value}))} style={{ width: "100%", padding: "5px 8px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12, boxSizing: "border-box" }} placeholder="e.g. Pixl Global" />
        </div>}
        {form.askType === "general_role" && <div style={{ gridColumn: "span 2" }}>
          <label style={{ fontSize: 10, fontWeight: 600 }}>Role / Description</label>
          <input value={form.targetRole} onChange={e => setForm(p => ({...p, targetRole: e.target.value}))} style={{ width: "100%", padding: "5px 8px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12, boxSizing: "border-box" }} placeholder="e.g. CFOs of companies for ERP solutions" />
        </div>}
        <div>
          <label style={{ fontSize: 10, fontWeight: 600 }}>Relevant Category</label>
          <select value={form.targetCategory} onChange={e => setForm(p => ({...p, targetCategory: e.target.value}))} style={{ width: "100%", padding: "5px 8px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12 }}>
            <option value="">Select category...</option>
            {allCategories.map(c => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
        <div><label style={{ fontSize: 10, fontWeight: 600 }}>Notes</label><input value={form.notes} onChange={e => setForm(p => ({...p, notes: e.target.value}))} style={{ width: "100%", padding: "5px 8px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12, boxSizing: "border-box" }} /></div>
      </div>
      <button onClick={addAsk} style={{ marginTop: 10, background: "#059669", color: "#fff", border: "none", borderRadius: 8, padding: "7px 16px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>Save Ask</button>
    </Card>}

    {asks.map(a => <Card key={a.id} style={{ marginBottom: 8, padding: 12, opacity: a.status === "fulfilled" ? 0.5 : 1 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div style={{ flex: 1 }}>
          <div style={{ fontWeight: 700, fontSize: 13 }}>{a.memberName}</div>
          <div style={{ fontSize: 12, color: "#374151", marginTop: 2 }}>
            {a.askType === "specific_person" && <span>🔍 Looking for: <strong>{a.targetName}</strong> {a.targetCompany && `from ${a.targetCompany}`}</span>}
            {a.askType === "specific_company" && <span>🏢 Looking for someone from: <strong>{a.targetCompany}</strong></span>}
            {a.askType === "general_role" && <span>👤 Looking for: <strong>{a.targetRole}</strong></span>}
            {a.askType === "free_text" && <span>🎯 Ask: <strong>{a.notes}</strong></span>}
          </div>
          <div style={{ fontSize: 11, color: "#6B7280", marginTop: 2 }}>{[a.targetCategory, a.askType !== "free_text" ? a.notes : "", a.date].filter(Boolean).join(" • ")}</div>
        </div>
        <div style={{ display: "flex", gap: 6, alignItems: "center", flexShrink: 0 }}>
          <Badge bg={isArchivedAsk(a) ? "#E5E7EB" : a.status === "open" ? "#FEF3C7" : "#D1FAE5"} text={isArchivedAsk(a) ? "#374151" : a.status === "open" ? "#92400E" : "#065F46"} label={isArchivedAsk(a) ? `🗄 Archived (${askAgeDays(a.date)}d old)` : a.status === "open" ? "Open" : "Fulfilled"} />
          {a.status === "open" && <button onClick={() => closeAsk(a.id)} style={{ fontSize: 10, padding: "3px 8px", border: "1px solid #D1D5DB", borderRadius: 4, cursor: "pointer", background: "#fff" }}>✓ Close</button>}
          <button onClick={() => setConfirmDelete(a)} style={{ fontSize: 10, padding: "3px 8px", border: "1px solid #FCA5A5", borderRadius: 4, cursor: "pointer", background: "#fff", color: "#991B1B" }}>🗑️</button>
        </div>
      </div>
    </Card>)}
  </div>;
}

function AIMatchTab({ visitors, asks, members }) {
  const [selectedVisitor, setSelectedVisitor] = useState(null);
  const matches = selectedVisitor ? findMatches(selectedVisitor, asks, members) : [];

  return <div>
    <Card style={{ marginBottom: 12, background: "#F5F3FF", borderColor: "#8B5CF6" }}>
      <div style={{ fontSize: 12, fontWeight: 700, color: "#5B21B6", marginBottom: 4 }}>🤖 AI Visitor-Member Matching</div>
      <div style={{ fontSize: 11, color: "#6D28D9", lineHeight: 1.6 }}>
        Select a visitor to see which members they should meet based on open asks, Contact Sphere category matches, and specialty relevance.
      </div>
    </Card>

    <div style={{ fontSize: 12, fontWeight: 600, marginBottom: 8 }}>Select a visitor:</div>
    <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 16 }}>
      {visitors.map(v => (
        <button key={v.id} onClick={() => setSelectedVisitor(v)}
          style={{ padding: "6px 12px", borderRadius: 8, fontSize: 12, cursor: "pointer", fontWeight: selectedVisitor?.id === v.id ? 700 : 400,
            background: selectedVisitor?.id === v.id ? "#8B1A1A" : "#fff",
            color: selectedVisitor?.id === v.id ? "#fff" : "#374151",
            border: `1px solid ${selectedVisitor?.id === v.id ? "#8B1A1A" : "#D1D5DB"}` }}>
          {v.name}
        </button>
      ))}
    </div>

    {selectedVisitor && <div>
      <Card style={{ marginBottom: 12, background: "#FFFBEB", borderColor: "#F59E0B" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <div style={{ fontWeight: 700, fontSize: 14 }}>{selectedVisitor.name}</div>
            <div style={{ fontSize: 12, color: "#92400E" }}>{selectedVisitor.business}</div>
            <div style={{ fontSize: 11, color: "#B45309" }}>{selectedVisitor.category} • {selectedVisitor.specialty}</div>
          </div>
          <StatusBadge status={selectedVisitor.status} />
        </div>
      </Card>

      <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 8 }}>🎯 Matches Found: {matches.length}</div>
      {matches.length === 0 && <div style={{ fontSize: 12, color: "#9CA3AF" }}>No matches found.</div>}
      {matches.map((m, i) => (
        <Card key={i} style={{ marginBottom: 8, padding: 12, borderLeft: `4px solid ${m.score >= 90 ? "#EF4444" : m.score >= 70 ? "#F59E0B" : "#3B82F6"}` }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4 }}>
            <div style={{ fontWeight: 700, fontSize: 13 }}>
              {m.member?.name}
              <span style={{ fontSize: 11, fontWeight: 400, color: "#6B7280", marginLeft: 6 }}>{m.member?.specialty}</span>
            </div>
            <span style={{ background: m.score >= 90 ? "#FEE2E2" : m.score >= 70 ? "#FEF3C7" : "#DBEAFE", color: m.score >= 90 ? "#991B1B" : m.score >= 70 ? "#92400E" : "#1E40AF", padding: "2px 10px", borderRadius: 20, fontSize: 11, fontWeight: 700 }}>{m.score}% match</span>
          </div>
          <div style={{ fontSize: 12, color: "#374151" }}>{m.reason}</div>
          {m.type === "ask" && <div style={{ fontSize: 10, color: "#6B7280", marginTop: 4 }}>From ask: {m.ask.notes} ({m.ask.date})</div>}
          <div style={{ marginTop: 6 }}>
            <Badge bg="#EEF2FF" text="#4338CA" label={m.type === "ask" ? "📌 Open Ask Match" : "🔄 Contact Sphere"} />
          </div>
        </Card>
      ))}
      {matches.length > 0 && <Card style={{ marginTop: 12, background: "#F0FDF4", borderColor: "#22C55E" }}>
        <div style={{ fontSize: 12, fontWeight: 700, color: "#15803D", marginBottom: 4 }}>💡 Suggested Actions</div>
        <div style={{ fontSize: 11, color: "#166534", lineHeight: 1.7 }}>
          {matches.filter(m => m.score >= 70).length > 0 && <>• Inform <strong>{matches.filter(m => m.score >= 70).map(m => m.member?.name).join(", ")}</strong> about this visitor before the meeting<br/></>}
          • Seat {selectedVisitor.name} next to: <strong>{matches[0]?.member?.name}</strong><br/>
          {matches.length > 1 && <>• Also introduce to: <strong>{matches.slice(1, 3).map(m => m.member?.name).join(", ")}</strong> during open networking</>}
        </div>
      </Card>}
    </div>}
  </div>;
}

// ═══════════════════════════════════════════
// UPGRADED: MEMBERS TAB — add, edit, delete members
// ═══════════════════════════════════════════
function MembersTab({ members, setMembers }) {
  const [search, setSearch] = useState("");
  const [filterCat, setFilterCat] = useState("");
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [form, setForm] = useState({ name: "", category: "", specialty: "" });
  const [confirmDelete, setConfirmDelete] = useState(null);

  // Build a category list that includes both existing chapter categories AND the master BNI list
  const existingCategories = [...new Set(members.map(m => m.category))];
  const categoryOptions = [...new Set([...existingCategories, ...ALL_BNI_CATEGORIES])].sort();

  const filtered = members.filter(m => {
    const matchSearch = !search || m.name.toLowerCase().includes(search.toLowerCase()) || m.specialty.toLowerCase().includes(search.toLowerCase());
    const matchCat = !filterCat || m.category === filterCat;
    return matchSearch && matchCat;
  });

  const addMember = () => {
    if (!form.name.trim() || !form.category || !form.specialty.trim()) {
      alert("Please fill in name, category, and specialty.");
      return;
    }
    const newId = members.length > 0 ? Math.max(...members.map(m => m.id)) + 1 : 1;
    setMembers(p => [...p, { id: newId, name: form.name.trim(), category: form.category, specialty: form.specialty.trim() }]);
    setForm({ name: "", category: "", specialty: "" });
    setShowForm(false);
  };

  const startEdit = (m) => { setEditingId(m.id); setForm({ name: m.name, category: m.category, specialty: m.specialty }); setShowForm(false); };
  const cancelEdit = () => { setEditingId(null); setForm({ name: "", category: "", specialty: "" }); };
  const saveEdit = () => {
    if (!form.name.trim() || !form.category || !form.specialty.trim()) {
      alert("Please fill in name, category, and specialty.");
      return;
    }
    setMembers(p => p.map(m => m.id === editingId ? { ...m, name: form.name.trim(), category: form.category, specialty: form.specialty.trim() } : m));
    setEditingId(null);
    setForm({ name: "", category: "", specialty: "" });
  };

  const doDelete = () => { setMembers(p => p.filter(m => m.id !== confirmDelete.id)); setConfirmDelete(null); };

  return (
    <div>
      <ConfirmModal
        open={!!confirmDelete}
        title="Remove this member from the chapter?"
        message={confirmDelete ? `"${confirmDelete.name}" (${confirmDelete.specialty}) will be removed from your member roster. This cannot be undone. Their existing asks and historical data are kept but will no longer match in AI Match.` : ""}
        confirmLabel="Yes, remove"
        onConfirm={doDelete}
        onCancel={() => setConfirmDelete(null)}
      />

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 8 }}>
        <span style={{ fontWeight: 700, fontSize: 15 }}>Chapter Members ({members.length})</span>
        <button
          onClick={() => { setShowForm(!showForm); cancelEdit(); }}
          style={{ background: "#8B1A1A", color: "#fff", border: "none", borderRadius: 8, padding: "7px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
          + Add Member
        </button>
      </div>

      {(showForm || editingId) && (
        <Card style={{ marginBottom: 12, background: editingId ? "#FEFCE8" : "#F0FDF4", borderColor: editingId ? "#F59E0B" : "#22C55E" }}>
          <div style={{ fontSize: 12, fontWeight: 800, color: editingId ? "#92400E" : "#15803D", marginBottom: 8 }}>
            {editingId ? "✏️ Editing member" : "➕ New member"}
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
            <div style={{ gridColumn: "span 2" }}>
              <label style={{ fontSize: 10, fontWeight: 600 }}>Full Name</label>
              <input value={form.name} onChange={e => setForm(p => ({...p, name: e.target.value}))} style={{ width: "100%", padding: "6px 10px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12, boxSizing: "border-box" }} placeholder="e.g. John Smith" />
            </div>
            <div>
              <label style={{ fontSize: 10, fontWeight: 600 }}>Category</label>
              <select value={form.category} onChange={e => setForm(p => ({...p, category: e.target.value}))} style={{ width: "100%", padding: "6px 10px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12 }}>
                <option value="">Select category...</option>
                {categoryOptions.map(c => <option key={c} value={c}>{c}</option>)}
              </select>
            </div>
            <div>
              <label style={{ fontSize: 10, fontWeight: 600 }}>Specialty / Classification</label>
              <input value={form.specialty} onChange={e => setForm(p => ({...p, specialty: e.target.value}))} style={{ width: "100%", padding: "6px 10px", border: "1px solid #D1D5DB", borderRadius: 6, fontSize: 12, boxSizing: "border-box" }} placeholder="e.g. Wealth Management" />
            </div>
          </div>
          <div style={{ display: "flex", gap: 6, marginTop: 10 }}>
            {editingId ? (
              <>
                <button onClick={saveEdit} style={{ background: "#059669", color: "#fff", border: "none", borderRadius: 8, padding: "7px 16px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>💾 Save Changes</button>
                <button onClick={cancelEdit} style={{ background: "#fff", color: "#374151", border: "1px solid #D1D5DB", borderRadius: 8, padding: "7px 16px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>Cancel</button>
              </>
            ) : (
              <>
                <button onClick={addMember} style={{ background: "#059669", color: "#fff", border: "none", borderRadius: 8, padding: "7px 16px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>✓ Add to Chapter</button>
                <button onClick={() => { setShowForm(false); setForm({ name: "", category: "", specialty: "" }); }} style={{ background: "#fff", color: "#374151", border: "1px solid #D1D5DB", borderRadius: 8, padding: "7px 16px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>Cancel</button>
              </>
            )}
          </div>
        </Card>
      )}

      <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
        <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search members..." style={{ flex: 1, padding: "7px 12px", border: "1px solid #D1D5DB", borderRadius: 8, fontSize: 12 }} />
        <select value={filterCat} onChange={e => setFilterCat(e.target.value)} style={{ padding: "7px 8px", border: "1px solid #D1D5DB", borderRadius: 8, fontSize: 12 }}>
          <option value="">All Categories</option>
          {existingCategories.sort().map(c => <option key={c} value={c}>{c}</option>)}
        </select>
      </div>
      <div style={{ fontSize: 12, color: "#6B7280", marginBottom: 8 }}>{filtered.length} members shown</div>
      <div style={{ maxHeight: 500, overflowY: "auto", border: "1px solid #E5E7EB", borderRadius: 10, background: "#fff" }}>
        {filtered.map(m => (
          <div key={m.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 12px", borderBottom: "1px solid #F3F4F6", fontSize: 12 }}>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontWeight: 700, color: "#111" }}>{m.name}</div>
              <div style={{ fontSize: 11, color: "#6B7280" }}>{m.specialty}</div>
            </div>
            <div style={{ display: "flex", gap: 6, alignItems: "center", flexShrink: 0 }}>
              <Badge bg="#F3F4F6" text="#374151" label={m.category} />
              <button onClick={() => startEdit(m)} style={{ background: "#FEF3C7", border: "1px solid #FCD34D", color: "#92400E", borderRadius: 6, padding: "3px 8px", fontSize: 10, cursor: "pointer", fontWeight: 700 }}>✏️</button>
              <button onClick={() => setConfirmDelete(m)} style={{ background: "#fff", border: "1px solid #FCA5A5", color: "#991B1B", borderRadius: 6, padding: "3px 8px", fontSize: 10, cursor: "pointer", fontWeight: 700 }}>🗑️</button>
            </div>
          </div>
        ))}
        {filtered.length === 0 && <div style={{ padding: 20, textAlign: "center", color: "#9CA3AF", fontSize: 12 }}>No members match your filter.</div>}
      </div>
    </div>
  );
}

function FollowUpTab({ visitors, setVisitors }) {
  const needsFollowUp = visitors.filter(v => ["attended","oriented"].includes(v.status));
  const genMsg = (v, type) => {
    const first = v.name.split(" ")[0];
    const msgs = {
      thankYou: `Hi ${first}, it was wonderful having you at BNI Insomniacs today!\n\nI hope you enjoyed meeting our members. I'd love to set up a 1-2-1 this week to learn more about your ${v.business} and see how I can refer people your way.\n\nWould tomorrow or Wednesday work?\n\n— [Your Name]`,
      ready: `That's fantastic, ${first}! Here's the application link: [link]\n\nHappy to walk you through it. Welcome to BNI Insomniacs!`,
      questions: `Completely understand, ${first}. I'd love to share how BNI Insomniacs has helped members grow their business. Can we grab a quick call?\n\nOur President would also love to connect.`,
      notNow: `No problem at all, ${first}. Thank you for visiting BNI Insomniacs.\n\nIs there any member you'd like me to connect you with? You're always welcome to visit again.`,
    };
    return msgs[type] || "";
  };
  const [msg, setMsg] = useState("");
  const [activeId, setActiveId] = useState(null);

  return <div>
    <Card style={{ marginBottom: 12, background: "#EEF2FF", borderColor: "#818CF8" }}>
      <div style={{ fontSize: 12, fontWeight: 700, color: "#4338CA", marginBottom: 4 }}>3-Response Follow-Up Framework</div>
      <div style={{ fontSize: 11, color: "#3730A3", lineHeight: 1.6 }}>
        <strong>1. Ready to apply</strong> → Send application link | <strong>2. More questions</strong> → Clarify with stories | <strong>3. Not now</strong> → Thank, offer connections
      </div>
    </Card>
    <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 8 }}>Needs Follow-Up ({needsFollowUp.length})</div>
    {needsFollowUp.length === 0 && <div style={{ color: "#9CA3AF", fontSize: 12 }}>All caught up! ✅</div>}
    {needsFollowUp.map(v => <Card key={v.id} style={{ marginBottom: 10, padding: 12 }}>
      <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 4 }}>{v.name} <span style={{ fontWeight: 400, fontSize: 11, color: "#6B7280" }}>{v.business}</span></div>
      <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 6 }}>
        {[["thankYou","💌 Thank You"],["ready","✅ Ready to Apply"],["questions","❓ More Questions"],["notNow","👋 Not Now"]].map(([k,l]) => (
          <button key={k} onClick={() => { setMsg(genMsg(v, k)); setActiveId(v.id); }} style={{ padding: "5px 10px", fontSize: 11, borderRadius: 6, border: "1px solid #D1D5DB", background: "#fff", cursor: "pointer", fontWeight: 600 }}>{l}</button>
        ))}
      </div>
      {activeId === v.id && msg && <div style={{ background: "#F5F3FF", border: "1px solid #DDD6FE", borderRadius: 8, padding: 10, marginTop: 6 }}>
        <pre style={{ fontSize: 11, whiteSpace: "pre-wrap", margin: 0, fontFamily: "inherit", lineHeight: 1.5, color: "#374151" }}>{msg}</pre>
        <button onClick={() => navigator.clipboard?.writeText(msg)} style={{ marginTop: 6, background: "#4338CA", color: "#fff", border: "none", borderRadius: 6, padding: "4px 12px", fontSize: 10, cursor: "pointer", fontWeight: 600 }}>📋 Copy</button>
      </div>}
    </Card>)}
  </div>;
}

// ═══════════════════════════════════════════
// CONNECTION ENGINE — ask aging, AI weekly digest, member 1-2-1 pairings
// ═══════════════════════════════════════════
function ConnectionEngineTab({ visitors, asks, members }) {
  const [digest, setDigest] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [copied, setCopied] = useState(false);

  const openAsks = asks.filter(a => a.status === "open");
  const activeAsks = openAsks.filter(isActiveAsk);
  const archivedAsks = openAsks.filter(isArchivedAsk);
  const [showArchived, setShowArchived] = useState(false);
  const daysOpen = (d) => askAgeDays(d);
  const agedAsks = [...activeAsks].sort((a, b) => daysOpen(b.date) - daysOpen(a.date));
  const thisWeek = visitors.filter(v => v.date === MEETING_DATE);

  const askSummary = (a) => [a.targetName, a.targetCompany, a.targetRole, a.notes].filter(Boolean).join(" — ");

  const generate = async () => {
    setLoading(true); setError(null); setCopied(false);
    try {
      const prompt = `You are the connection strategist for BNI Insomniacs, a BNI chapter in Dubai. The Visitor Host needs this week's connection digest before the Wednesday meeting on ${MEETING_DATE}.

THIS WEEK'S VISITORS (${thisWeek.length}):
${thisWeek.length ? thisWeek.map(v => `- ${v.name} | ${v.business || "no company"} | ${v.category || "no category"} / ${v.specialty || ""} | status: ${v.status} | invited by: ${v.invitedBy || "unknown"} | notes: ${v.callNotes || "none"}`).join("\n") : "(none registered yet)"}

ACTIVE OPEN ASKS — last 6 weeks (${activeAsks.length}):
${activeAsks.map(a => `- ${a.memberName} (open ${daysOpen(a.date)} days): wants ${askSummary(a) || a.targetCategory}`).join("\n") || "(none)"}

ARCHIVED ASKS — older than 6 weeks, no longer actively pursued (${archivedAsks.length}):
${archivedAsks.map(a => `- ${a.memberName} (${daysOpen(a.date)} days old): wanted ${askSummary(a) || a.targetCategory}`).join("\n") || "(none)"}

CHAPTER MEMBERS (${members.length}):
${members.map(m => `${m.name} — ${m.category} / ${m.specialty}`).join("\n")}

Think like a master networker: match visitors to asks AND to members who would naturally refer business to each other (contact spheres, supply chains, shared client types — connections can cross categories). Also pair members whose open asks or specialties complement each other for 1-2-1 meetings.

ARCHIVED ASK RULE: match visitors primarily against ACTIVE asks. But also scan the ARCHIVED asks — if this week's visitor fits an archived ask, or even comes close, that is gold: flag it as a revival so the Visitor Host can tell the member their old ask just walked through the door. Only include genuine fits.

STRICT RULES:
- Only use names that appear in the data above. Never invent people.
- If there are no visitors this week, focus on member 1-2-1 pairings and ask insights.
- Respond with ONLY a JSON object, no markdown fences, no preamble, in exactly this shape:
{
  "headline": "One energising sentence summarising this week's biggest connection opportunity",
  "visitorMatches": [{ "visitorName": "...", "members": ["member name", "member name"], "reason": "specific, concrete reason (max 30 words)" }],
  "archivedAskRevivals": [{ "visitorName": "...", "memberName": "...", "askSummary": "what the member originally asked for", "reason": "why this visitor fits or comes close (max 25 words)" }],
  "oneToOnes": [{ "memberA": "...", "memberB": "...", "reason": "why this 1-2-1 makes business sense right now (max 30 words)" }],
  "askInsights": ["short observation about the open asks, e.g. which are going stale or which categories dominate (max 3 items)"],
  "whatsappMessage": "A short, friendly pre-meeting message (with a couple of emoji) for the chapter leadership WhatsApp group summarising the top visitor-member connections, any archived-ask revival, and one suggested 1-2-1. Plain text, under 120 words."
}
archivedAskRevivals must be an empty array if nothing genuinely fits. Include every visitor in visitorMatches (best-effort matches, empty members array if truly nothing fits). Give 3 to 5 oneToOnes.`;

      const parsed = await callClaude(prompt, 2000);
      setDigest(parsed);
    } catch (e) {
      console.error("Digest generation failed:", e);
      setError("Could not generate the digest. Check the API key in Vercel and try again.");
    }
    setLoading(false);
  };

  const copyWhatsApp = () => {
    navigator.clipboard?.writeText(digest.whatsappMessage);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const ageBadge = (days) => days >= 35 ? { bg: "#FEE2E2", text: "#991B1B", label: `🔴 ${days}d — archives at 42d` } :
                            days >= 21 ? { bg: "#FEF3C7", text: "#92400E", label: `🟡 ${days}d open` } :
                            { bg: "#DBEAFE", text: "#1E40AF", label: `${days}d open` };

  return <div>
    <Card style={{ marginBottom: 12, background: "#F5F3FF", borderColor: "#8B5CF6" }}>
      <div style={{ fontSize: 12, fontWeight: 700, color: "#5B21B6", marginBottom: 4 }}>🔗 Connection Engine</div>
      <div style={{ fontSize: 11, color: "#6D28D9", lineHeight: 1.6 }}>
        Turns your weekly asks data into action: which visitors can close open asks, which members should book a 1-2-1, and which asks are going stale — plus a ready-to-send WhatsApp digest for chapter leadership. Asks auto-archive after 6 weeks but the AI still watches them for visitor matches.
      </div>
    </Card>

    {/* Ask aging — instant, no AI needed */}
    <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 8 }}>⏳ Active Ask Ageing ({activeAsks.length})</div>
    {activeAsks.length === 0 && <div style={{ fontSize: 12, color: "#9CA3AF", marginBottom: 16 }}>No active asks. Capture this week's asks in the Asks tab to power the digest.</div>}
    {agedAsks.map(a => {
      const d = daysOpen(a.date);
      const b = ageBadge(d);
      return (
        <Card key={a.id} style={{ marginBottom: 8, padding: 12, borderLeft: `4px solid ${d >= 35 ? "#EF4444" : d >= 21 ? "#F59E0B" : "#3B82F6"}` }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontWeight: 700, fontSize: 13 }}>{a.memberName}</div>
              <div style={{ fontSize: 12, color: "#374151" }}>{askSummary(a) || a.targetCategory}</div>
              <div style={{ fontSize: 10, color: "#6B7280", marginTop: 2 }}>Since {a.date}</div>
            </div>
            <span style={{ background: b.bg, color: b.text, padding: "2px 10px", borderRadius: 20, fontSize: 10, fontWeight: 700, whiteSpace: "nowrap" }}>{b.label}</span>
          </div>
          {d >= 35 && <div style={{ fontSize: 11, color: "#991B1B", marginTop: 6, background: "#FEF2F2", borderRadius: 6, padding: "5px 8px" }}>💬 Suggest {a.memberName.split(" ")[0]} refreshes or re-presents this ask — it auto-archives at 6 weeks.</div>}
        </Card>
      );
    })}

    {/* Archived asks — collapsed, still feed the AI */}
    {archivedAsks.length > 0 && (
      <div style={{ marginBottom: 4 }}>
        <button onClick={() => setShowArchived(!showArchived)} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 12, fontWeight: 700, color: "#6B7280", padding: "6px 0" }}>
          {showArchived ? "▾" : "▸"} 🗄 Archived asks ({archivedAsks.length}) — older than 6 weeks, still watched by the AI
        </button>
        {showArchived && archivedAsks.map(a => (
          <Card key={a.id} style={{ marginBottom: 6, padding: 10, background: "#F9FAFB", opacity: 0.75 }}>
            <div style={{ fontSize: 12 }}><b>{a.memberName}</b> <span style={{ color: "#6B7280" }}>— {askSummary(a) || a.targetCategory}</span></div>
            <div style={{ fontSize: 10, color: "#9CA3AF" }}>{daysOpen(a.date)} days old · since {a.date}</div>
          </Card>
        ))}
      </div>
    )}

    {/* AI digest */}
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", margin: "18px 0 8px" }}>
      <div style={{ fontWeight: 700, fontSize: 14 }}>🤖 Weekly Connection Digest</div>
      <button onClick={generate} disabled={loading || (openAsks.length === 0 && thisWeek.length === 0)}
        style={{ background: "linear-gradient(135deg, #7C3AED, #5B21B6)", color: "#fff", border: "none", borderRadius: 8, padding: "8px 16px", fontSize: 12, fontWeight: 700, cursor: loading ? "wait" : "pointer", opacity: (loading || (openAsks.length === 0 && thisWeek.length === 0)) ? 0.6 : 1 }}>
        {loading ? "Thinking…" : digest ? "↺ Regenerate" : "✨ Generate Digest"}
      </button>
    </div>
    {openAsks.length === 0 && thisWeek.length === 0 && <div style={{ fontSize: 11, color: "#9CA3AF" }}>Add this week's visitors or open asks first — the digest needs data to work with.</div>}
    {loading && (
      <div style={{ display: "flex", alignItems: "center", gap: 8, color: "#5B21B6", fontSize: 12, padding: "8px 0" }}>
        <div style={{ width: 16, height: 16, border: "2px solid #7C3AED", borderTopColor: "transparent", borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
        Analysing {thisWeek.length} visitors, {openAsks.length} open asks and {members.length} members…
      </div>
    )}
    {error && <div style={{ color: "#991B1B", fontSize: 12, marginBottom: 8 }}>{error}</div>}

    {digest && !loading && <div>
      <Card style={{ marginBottom: 12, background: "#FFFBEB", borderColor: "#F59E0B" }}>
        <div style={{ fontSize: 13, fontWeight: 800, color: "#92400E", lineHeight: 1.5 }}>💡 {digest.headline}</div>
      </Card>

      {digest.visitorMatches?.length > 0 && <>
        <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 8 }}>👥 Visitor → Member Connections</div>
        {digest.visitorMatches.map((vm, i) => (
          <Card key={i} style={{ marginBottom: 8, padding: 12, borderLeft: "4px solid #7C3AED" }}>
            <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 2 }}>{vm.visitorName}
              {vm.members?.length > 0 && <span style={{ fontWeight: 400, color: "#6B7280" }}> → meet </span>}
              {vm.members?.length > 0 && <span style={{ color: "#5B21B6" }}>{vm.members.join(", ")}</span>}
            </div>
            <div style={{ fontSize: 12, color: "#374151" }}>{vm.members?.length ? vm.reason : "No strong match this week — introduce during open networking."}</div>
          </Card>
        ))}
      </>}

      {digest.archivedAskRevivals?.length > 0 && <>
        <div style={{ fontWeight: 700, fontSize: 13, margin: "14px 0 8px" }}>🔁 Archived Ask Revivals</div>
        {digest.archivedAskRevivals.map((rv, i) => (
          <Card key={i} style={{ marginBottom: 8, padding: 12, borderLeft: "4px solid #D97706", background: "#FFFBEB" }}>
            <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 2 }}>{rv.visitorName} <span style={{ color: "#D97706" }}>fits</span> {rv.memberName}'s old ask</div>
            <div style={{ fontSize: 11, color: "#92400E", marginBottom: 3 }}>Original ask: {rv.askSummary}</div>
            <div style={{ fontSize: 12, color: "#374151" }}>{rv.reason}</div>
            <div style={{ fontSize: 10, color: "#B45309", marginTop: 4 }}>💬 Tell {(rv.memberName || "").split(" ")[0]} their ask just walked through the door.</div>
          </Card>
        ))}
      </>}

      {digest.oneToOnes?.length > 0 && <>
        <div style={{ fontWeight: 700, fontSize: 13, margin: "14px 0 8px" }}>🤝 Suggested Member 1-2-1s</div>
        {digest.oneToOnes.map((p, i) => (
          <Card key={i} style={{ marginBottom: 8, padding: 12, borderLeft: "4px solid #059669" }}>
            <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 2 }}>{p.memberA} <span style={{ color: "#059669" }}>⇄</span> {p.memberB}</div>
            <div style={{ fontSize: 12, color: "#374151" }}>{p.reason}</div>
          </Card>
        ))}
      </>}

      {digest.askInsights?.length > 0 && (
        <Card style={{ margin: "14px 0 12px", background: "#F0F9FF", borderColor: "#38BDF8" }}>
          <div style={{ fontSize: 12, fontWeight: 700, color: "#0C4A6E", marginBottom: 6 }}>📈 Ask Insights</div>
          {digest.askInsights.map((ins, i) => (
            <div key={i} style={{ fontSize: 11, color: "#0369A1", lineHeight: 1.6, marginBottom: 3 }}>• {ins}</div>
          ))}
        </Card>
      )}

      {digest.whatsappMessage && (
        <Card style={{ background: "#F0FDF4", borderColor: "#22C55E" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: "#15803D" }}>📱 Leadership WhatsApp Digest</div>
            <button onClick={copyWhatsApp} style={{ background: copied ? "#059669" : "#25D366", color: "#fff", border: "none", borderRadius: 8, padding: "6px 14px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>{copied ? "✓ Copied!" : "📋 Copy for WhatsApp"}</button>
          </div>
          <pre style={{ fontSize: 12, whiteSpace: "pre-wrap", margin: 0, fontFamily: "inherit", lineHeight: 1.6, color: "#166534", background: "#fff", borderRadius: 8, padding: 10, border: "1px solid #BBF7D0" }}>{digest.whatsappMessage}</pre>
        </Card>
      )}
    </div>}
  </div>;
}

export default function App() {
  const [tab, setTab] = useState(0);
  const [visitors, setVisitors] = useState(INITIAL_VISITORS);
  const [asks, setAsks] = useState(INITIAL_ASKS);
  const [members, setMembers] = useState(INITIAL_MEMBERS);
  const [archived, setArchived] = useState([]);
  const [loading, setLoading] = useState(true);

  // Load real data from Supabase on mount
  useEffect(() => {
    const loadData = async () => {
      try {
        const [{ data: visitorsData }, { data: asksData }] = await Promise.all([
          supabase.from("visitors").select("*").order("id", { ascending: true }),
          supabase.from("asks").select("*").order("id", { ascending: true }),
        ]);
        if (visitorsData?.length) {
          setVisitors(visitorsData.map(v => ({
            ...v,
            invitedBy: v.invited_by,
            callNotes: v.call_notes,
            seatAssignment: v.seat_assignment,
            followUpResponse: v.follow_up_response,
          })));
        }
        if (asksData?.length) {
          setAsks(asksData.map(a => ({
            ...a,
            memberId: a.member_id,
            memberName: a.member_name,
            askType: a.ask_type,
            targetName: a.target_name,
            targetCompany: a.target_company,
            targetCategory: a.target_category,
            targetRole: a.target_role,
          })));
        }
      } catch (e) {
        console.error("Failed to load from Supabase:", e);
      }
      setLoading(false);
    };
    loadData();
  }, []);

  if (loading) return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "center", minHeight: "100vh", background: "#F9FAFB", flexDirection: "column", gap: 12 }}>
      <div style={{ width: 32, height: 32, border: "3px solid #8B1A1A", borderTopColor: "transparent", borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
      <div style={{ fontSize: 13, color: "#6B7280", fontWeight: 600 }}>Loading BNI Insomniacs...</div>
      <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
    </div>
  );

  const tabs = [
    <DashboardTab visitors={visitors} asks={asks} members={members} archived={archived} />,
    <VisitorsTab visitors={visitors} setVisitors={setVisitors} asks={asks} members={members} archived={archived} setArchived={setArchived} />,
    <AsksTab asks={asks} setAsks={setAsks} members={members} />,
    <ConnectionEngineTab visitors={visitors} asks={asks} members={members} />,
    <AIMatchTab visitors={visitors} asks={asks} members={members} />,
    <SeatPlanner visitors={visitors} asks={asks} members={members} />,
    <MembersTab members={members} setMembers={setMembers} />,
    <ArchiveTab archived={archived} setArchived={setArchived} visitors={visitors} setVisitors={setVisitors} />,
    <FollowUpTab visitors={visitors} setVisitors={setVisitors} />,
  ];

  return <div style={{ fontFamily: "'Segoe UI', -apple-system, sans-serif", background: "#F9FAFB", minHeight: "100vh" }}>
    <div style={{ background: "linear-gradient(135deg, #8B1A1A 0%, #1B2A4A 100%)", padding: "14px 20px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
      <div>
        <div style={{ color: "#fff", fontSize: 17, fontWeight: 800, letterSpacing: -0.5 }}>BNI Insomniacs <span style={{ fontSize: 10, fontWeight: 700, background: "rgba(255,255,255,0.2)", padding: "2px 6px", borderRadius: 8, marginLeft: 6, verticalAlign: "middle" }}>v6.6</span></div>
        <div style={{ color: "#FFD4D4", fontSize: 10 }}>Visitor Host Command Centre • {members.length} Members</div>
      </div>
      <div style={{ display: "flex", gap: 12, color: "#FFD4D4", fontSize: 11 }}>
        <span>👥 {visitors.filter(v => v.date === MEETING_DATE).length} this week</span>
        <span>🎯 {asks.filter(isActiveAsk).length} open asks</span>
        <span>🗄️ {archived.length} archived</span>
      </div>
    </div>

    <div style={{ background: "#fff", borderBottom: "1px solid #E5E7EB", display: "flex", overflowX: "auto", padding: "0 4px" }}>
      {TABS.map((t, i) => (
        <button key={i} onClick={() => setTab(i)} style={{
          padding: "10px 14px", border: "none", background: "none", cursor: "pointer",
          fontSize: 12, fontWeight: tab === i ? 700 : 500,
          color: tab === i ? "#8B1A1A" : "#6B7280",
          borderBottom: tab === i ? "3px solid #8B1A1A" : "3px solid transparent",
          whiteSpace: "nowrap",
        }}>{t.icon} {t.label}</button>
      ))}
    </div>

    <div style={{ padding: 16, maxWidth: 900, margin: "0 auto" }}>
      {tabs[tab]}
    </div>
  </div>;
}
