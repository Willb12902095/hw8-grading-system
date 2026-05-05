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

The grading system applies a code-only rubric when no chat history is used. Each criterion is evaluated using rule-based analysis implemented in the script.

### 4.1 Code Submission (25 points)
Evaluates code structure and implementation effort:
- Number of files submitted
- Total lines of code
- Presence of functions (regex detection of `def`)
- Presence of classes (`class` keyword)
- Presence of comments and docstrings

Scoring logic:
- Higher scores for modular code (multiple functions/classes)
- Lower scores for minimal or single-function scripts

### 4.2 Completeness (25 points)
Checks whether required components are present:
- At least one valid code file detected → full score
- No code files → zero score

This criterion ensures the submission is not empty or invalid.

### 4.3 Code Correctness (35 points)
Evaluates logical structure using static pattern detection:
- Game loop detection (`while` loop)
- Pygame usage (`import pygame`)
- Rendering/update logic (`pygame.display`, `draw`, or `update`)
- Entry point detection (`if __name__ == "__main__"` or `main()`)

Scoring logic:
- Points added for each detected feature
- Points deducted if key components (e.g., game loop, entry point) are missing
- Syntax consistency checked via bracket/parentheses matching

### 4.4 Code Quality (15 points)
Evaluates readability and maintainability:
- Function modularity (number of functions)
- Use of classes
- Presence of comments and docstrings
- Naming conventions (variable/function patterns)
- Error handling (`try/except` blocks)

Higher scores are given to well-structured and readable code.

### 4.5 Runtime Behavior (10 points)
Evaluated using an optional runtime smoke test:
- Each Python file is executed for up to 5 seconds
- PASS → program runs without crashing
- PARTIAL → some files fail
- FAIL → program crashes

Scoring:
- PASS → 10 points
- PARTIAL → 5 points
- FAIL → 0 points

Note: This test checks execution stability only, not gameplay correctness.

### Total Score
Total = 110 points

Scores are converted into percentages and mapped to letter grades:
A (≥90), B (80–89), C (70–79), D (60–69), F (<60)

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