# Homework Grading System Report

## 1. Overview
This grading system automatically evaluates programming assignments submitted as a ZIP file. It applies a structured grading rubric while handling the complexity of real-world student submissions. Since submission formats are inconsistent, the system standardizes inputs without modifying original files.

## 2. Handling of Student Submission Structure
### 2.1 Motivation
Student submissions vary widely: single vs. multiple files, inconsistent naming, missing folders, and extra assets (e.g., images).

### 2.2 Organization Strategy
The system groups files using the filename prefix before “#”.
Example:
justincyc2025#_game.py → student: justincyc2025
This ensures correct grouping even without folders or consistent formatting.

### 2.3 Data Preservation
All files are kept as submitted. Only minimal cleanup (e.g., ignoring .DS_Store, __MACOSX) is performed. No code or structure is modified.

### 2.4 File Detection
Code files (.py, .java, .cpp) are detected by type, not location. Chat history files are also detected. Although present, chat histories were not used in this stage and will be analyzed separately.

## 3. Workflow
### 3.1 Input
python homework_grader.py [.zip]

### 3.2 Extraction
Files are extracted into a temporary directory and preserved in original form.

### 3.3 Static Analysis
Code is analyzed without execution, checking:
- structure and size
- functions and classes
- comments and naming
- key features (game loop, pygame, rendering, entry point)

### 3.4 Runtime Smoke Test
Optional 5-second execution per file:
RUN_SMOKE_TEST=1 python homework_grader.py [.zip]
This command activates the runtime smoke test ("smoking gun") to verify whether student programs can launch without crashing.
PASS – runs successfully
PARTIAL – mixed results
FAIL – crashes
This checks execution stability only.

## 4. Grading Rubric (Code-Only)
- Code Submission (25)
- Completeness (25)
- Code Correctness (35)
- Code Quality (15)
- Runtime Behavior (10)
Total: 110 points

## 5. Scoring
Grades are assigned by percentage:
A (≥90), B (80–89), C (70–79), D (60–69), F (<60)

## 6. Outputs
- HW8_final_grades.xlsx (scores + summary)
- HW8_final_grading_log.txt (detailed logs)

## 7. Results

A total of 9 students were evaluated.

- **HW 8-1 Average:** 88.1%  
  - High (≥80%): 8 students  
  - Medium (60–79%): 1 student  
  - Low (<60%): 0 students  

- **HW 8-2 Average:** 82.3%  
  - High (≥80%): 8 students  
  - Medium (60–79%): 0 students  
  - Low (<60%): 1 student  

Overall, most students achieved strong performance. The lower average in HW 8-2 is primarily due to runtime failures and a small number of incomplete or incorrect submissions, which significantly impacted overall scores.

Note: Detailed results for each student can be found in the generated grading log files.

## 8. Observations
Strengths:
- Most implemented game loops and pygame logic correctly
- Good structure and modularity in most of the submissions

Grader Observation (HW 8-2):

- Noticeable improvement in overall game interface and user experience

- Several students enhanced visual design by adding background images

- Some submissions included soundtracks or sound effects during gameplay

- These additions indicate better understanding of environment setup and asset integration

## 9. Limitations
- Runtime test does not evaluate gameplay quality
- Static analysis cannot guarantee correctness
- Missing external files may cause failures
- Some cases require manual review

## 10. Conclusion
The system provides efficient and consistent grading despite messy submissions. By combining file organization, rubric-based scoring, and runtime checks, it balances automation, fairness, and reliability.