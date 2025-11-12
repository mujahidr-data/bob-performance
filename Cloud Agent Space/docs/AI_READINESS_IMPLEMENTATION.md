# AI Readiness Implementation Guide

## Overview

This guide explains how to implement and use the AI Readiness Assessment feature in the Bob Performance Module.

---

## 🎯 Purpose

Identify employees who are ready to operate in an AI-first environment and gauge their potential for AI-augmented role transformation. This assessment helps leadership:

1. **Identify AI Champions**: Employees demonstrating AI-first behaviors
2. **Target Development**: Focus training on "AI-Capable" employees
3. **Strategic Planning**: Understand organizational AI readiness
4. **Performance Context**: Link AI adoption to performance outcomes

---

## 📊 Implementation Steps

### Step 1: Create AI Readiness Mapping Sheet

1. Open your Google Sheet with Bob Performance Module
2. Go to **🚀 Bob Performance Module** menu
3. Click **🎯 Create AI Readiness Mapping**
4. This creates a reference sheet mapping each department to its AI category

**What it does:**
- Maps departments (Engineering, Product, Sales, etc.) to AI transformation categories
- Provides category descriptions and assessment criteria
- Auto-populates in Summary sheet based on employee department

---

### Step 2: Build/Rebuild Summary Sheet

1. Ensure all source data is imported:
   - Base Data
   - Bonus History
   - Full Comp History
   - Performance Ratings
   - Bob Perf Report

2. Run **🔧 Build Summary Sheet**

3. The Summary sheet now includes a new column: **"AI Readiness Category"**
   - Auto-populated via VLOOKUP from AI Readiness Mapping
   - Based on employee's Department (column H)

---

### Step 3: Add Assessment Column (Manual)

Leadership needs to assess each employee. Add a new column after "AI Readiness Category":

1. **Column Name**: "AI Readiness Assessment"

2. **Data Validation** (Dropdown):
   - Select the entire column (e.g., column Z)
   - Go to **Data** → **Data validation**
   - Choose **List of items**
   - Enter: `AI-Ready, AI-Capable, Not AI-Ready, Not Assessed`
   - Check "Show dropdown list in cell"
   - Save

3. **Conditional Formatting**:
   - **AI-Ready**: Green (#b7d7a8)
   - **AI-Capable**: Amber (#fff2cc)
   - **Not AI-Ready**: Red (#f4c7c3)
   - **Not Assessed**: Grey (#efefef)

---

### Step 4: Leadership Assessment Process

For each employee, leaders should:

1. **Review AI Readiness Category**
   - Understand what the AI transformation means for that function
   - Reference: `docs/AI_READINESS_CATEGORIES.md`

2. **Evaluate Against Criteria**
   - Does the employee demonstrate 3+ of the 5 criteria? → **AI-Ready**
   - Does the employee demonstrate 2 criteria? → **AI-Capable**
   - Does the employee demonstrate 0-1 criteria? → **Not AI-Ready**

3. **Consider Performance Context**
   - High performers (HH, HM) should ideally be AI-Ready or AI-Capable
   - Low performers may have AI adoption as a development area

4. **Reference Manager Blurbs**
   - Check if AI usage is mentioned in manager feedback
   - Look for evidence of AI tool adoption, velocity improvements

---

## 📖 AI Readiness Criteria

Each function has 5 specific criteria. Here's a quick reference:

### Engineering: AI-Accelerated Developer
- ✅ Uses AI tools for 30%+ of development work
- ✅ Demonstrates 2x+ velocity improvement with AI assistance
- ✅ Builds automated testing/deployment pipelines with AI
- ✅ Shares AI development patterns with team
- ✅ Contributes to AI-driven code review processes

### Product: AI-Forward Solutions Engineer
- ✅ Uses AI to analyze client feedback and requirements
- ✅ Builds AI-powered prototypes/demos for client validation
- ✅ Leverages AI for competitive intelligence gathering
- ✅ Creates AI-assisted product specifications and roadmaps
- ✅ Demonstrates ability to customize solutions using AI tools

### Customer Success: AI Outcome Manager
- ✅ Uses AI for churn prediction and proactive intervention
- ✅ Leverages AI chatbots/agents for L1 support automation
- ✅ Creates AI-powered customer health dashboards
- ✅ Manages 2x+ customer portfolio with AI assistance
- ✅ Uses AI to generate personalized customer insights/recommendations

### Marketing: AI Content Intelligence Strategist
- ✅ Uses AI to generate and iterate marketing content
- ✅ Leverages AI for competitive intelligence synthesis
- ✅ Creates personalized content using AI segmentation
- ✅ Uses AI for campaign performance optimization
- ✅ Integrates insights from CS/support data via AI analysis

### Sales: AI-Augmented Revenue Accelerator
- ✅ Uses AI for lead prioritization and scoring
- ✅ Leverages AI for personalized outreach generation
- ✅ Automates meeting prep/follow-ups with AI
- ✅ Uses AI for deal risk assessment and forecasting
- ✅ Demonstrates increased pipeline velocity with AI tools

**For other functions, see:** `docs/AI_READINESS_CATEGORIES.md`

---

## 🎨 Example Summary Sheet Layout

| Emp ID | Emp Name | Department | ... | AI Readiness Category | AI Readiness Assessment |
|--------|----------|------------|-----|----------------------|------------------------|
| 12345 | John Doe | Engineering | ... | AI-Accelerated Developer | AI-Ready |
| 67890 | Jane Smith | Product | ... | AI-Forward Solutions Engineer | AI-Capable |
| 11111 | Bob Johnson | Sales | ... | AI-Augmented Revenue Accelerator | Not AI-Ready |

---

## 📈 Reporting & Analytics

### Key Metrics to Track

1. **AI Readiness Distribution**
   - % AI-Ready by function
   - % AI-Capable by function
   - % Not AI-Ready by function

2. **Correlation with Performance**
   - Do HH employees have higher AI-Ready rates?
   - Do NI employees have lower AI-Ready rates?

3. **Department Comparison**
   - Which functions are most AI-ready?
   - Where is training/coaching needed?

4. **Trend Over Time**
   - Track AI readiness progression across performance cycles
   - Measure impact of AI training programs

### Sample Pivot Tables

**AI Readiness by Department:**
```
Rows: Department
Columns: AI Readiness Assessment
Values: Count of Emp ID
```

**AI Readiness by Performance Rating:**
```
Rows: Q2/Q3 Rating
Columns: AI Readiness Assessment
Values: Count of Emp ID
```

---

## 🚀 Rollout Strategy

### Phase 1: Pilot (2 weeks)
- Select 2-3 departments for pilot assessment
- Train leadership on criteria
- Gather feedback, refine criteria if needed

### Phase 2: Calibration (1 week)
- Leadership reviews pilot assessments
- Calibrates understanding of "AI-Ready" vs "AI-Capable"
- Addresses edge cases

### Phase 3: Full Rollout (3 weeks)
- Assess all employees across all functions
- Create development plans for "AI-Capable" and "Not AI-Ready"
- Communicate expectations to employees

### Phase 4: Ongoing (Quarterly)
- Reassess AI readiness each performance cycle
- Track progression and impact of development programs
- Update criteria as AI capabilities evolve

---

## 💡 Best Practices

### For Leadership

1. **Be Specific**: Use concrete examples of AI tool usage
2. **Be Fair**: Don't penalize employees in roles with limited AI applicability
3. **Be Developmental**: Frame as growth opportunity, not just evaluation
4. **Be Consistent**: Calibrate across managers to ensure fairness

### For HR/People Ops

1. **Communicate Purpose**: Explain why AI readiness matters for company strategy
2. **Provide Resources**: Share AI tool training, documentation, examples
3. **Track Progress**: Monitor changes over time, celebrate improvements
4. **Link to Development**: Create clear paths from "Not AI-Ready" → "AI-Ready"

### For Employees

1. **Self-Assess**: Review criteria for your function before 1:1
2. **Show Evidence**: Bring examples of AI tool usage, productivity gains
3. **Ask for Help**: Request training or coaching if needed
4. **Experiment**: Try AI tools in low-risk scenarios first

---

## 🔧 Customization

### Adding New Functions

1. Open "AI Readiness Mapping" sheet
2. Add new row: `[Function Name, AI Category Name]`
3. Create criteria for new category in `AI_READINESS_CATEGORIES.md`

### Modifying Categories

1. Update category descriptions in `AI_READINESS_CATEGORIES.md`
2. Update "AI Readiness Mapping" sheet if category names change
3. Communicate changes to leadership
4. Re-run "Build Summary Sheet" to refresh

### Adding Custom Assessment Levels

Instead of 3 levels (AI-Ready, AI-Capable, Not AI-Ready), you could use:
- 5 levels: Advanced, Proficient, Developing, Beginner, Not Exposed
- Or role-specific labels

Just update the data validation dropdown in Step 3.

---

## 🆘 Troubleshooting

### "Not Mapped" appears in AI Readiness Category

**Cause**: Employee's department doesn't exist in AI Readiness Mapping sheet

**Fix**:
1. Check employee's department in Base Data
2. Add that department to AI Readiness Mapping sheet
3. Re-run "Build Summary Sheet"

### Category doesn't match employee's actual role

**Cause**: Department field is too broad or inaccurate

**Fix**:
1. Use a more specific department field (e.g., "Team" or "Sub-Department")
2. Modify VLOOKUP formula in buildSummarySheet to reference different column
3. Or manually override in Summary sheet

### Assessment dropdown not appearing

**Cause**: Data validation not set up correctly

**Fix**:
1. Ensure data validation is applied to entire column
2. Check that dropdown values are spelled exactly as specified
3. Remove and re-add data validation if needed

---

## 📚 Related Documentation

- **AI Readiness Categories**: `docs/AI_READINESS_CATEGORIES.md`
- **Manager Blurb Generator**: `docs/MANAGER_BLURB_GENERATOR.md`
- **Project README**: `README.md`

---

## 🔄 Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0 | 2025-11-12 | Initial implementation with 12 function categories |

---

**Questions?** Contact the HR/People Ops team or refer to `docs/AI_READINESS_CATEGORIES.md` for detailed criteria.

