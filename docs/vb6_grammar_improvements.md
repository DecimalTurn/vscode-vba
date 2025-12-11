# VB6 Grammar Improvements from VBA Grammar

This document outlines incremental improvements that can be made to the vb6 grammar by incorporating patterns from the VBA grammar. The improvements are organized by priority and complexity.

## Priority 1: Low-Hanging Fruit
*Easy wins with immediate benefit*

### 1. Improve String Literal Handling
- **Current**: Basic double-quote matching
- **VBA Enhancement**: Handles escaped quotes (`""`) and type suffixes properly
- **Benefit**: More accurate string highlighting, handles strings with embedded quotes
- **Pattern**: `'"("")*([^"\n]*)((?:"")[^"\n]+)?"("")*'`

### 2. Add Number Type Suffixes
- **Current**: Just the number value
- **VBA Enhancement**: Recognizes type hint characters (`%`, `&`, `@`, `!`, `#`, `$`)
- **Benefit**: Distinguishes integer%, long&, decimal@, single!, double#, currency$
- **Pattern**: `-?\d+\.?\d*[%&@!#]?`

### 3. Add Line Continuation Handling
- **Current**: Basic comment continuation only
- **VBA Enhancement**: Dedicated line continuation pattern with proper scope
- **Benefit**: Better multi-line statement handling
- **Pattern**: `(?<=\s)_\s*\n` with scope `keyword.control.line-continuation.vb6`

### 4. Add Line Separator (Colon)
- **Current**: Not explicitly handled
- **VBA Enhancement**: Highlights `:` as statement separator
- **Benefit**: Better readability for multi-statement lines
- **Scope**: `keyword.control.line-separator.vb6`

---

## Priority 2: Operators
*Moderately easy, high impact*

### 5. Categorize Operators
- **Current**: Operators mixed in keywords
- **VBA Enhancement**: Separate categories for different operator types:
  - **Arithmetic**: `*, /, \, ^, +, -, Mod`
  - **Comparison**: `<, >, =, <=, >=, <>, Is, Like`
  - **Logical**: `And, Or, Not, Xor, Eqv, Imp`
  - **Concatenation**: `&, +`
  - **New**: Object instantiation operator
- **Benefit**: More granular semantic highlighting

### 6. Add Range Operator
- **Current**: `To` is in keywords
- **VBA Enhancement**: Separate `To` as `keyword.operator.range.vb6`
- **Benefit**: Distinguishes range usage from other keywords

---

## Priority 3: Better Structure
*Medium complexity*

### 7. Structured Method/Property Signatures
- **Current**: Basic function/sub detection
- **VBA Enhancement**: Complete method signature parsing with:
  - Visibility modifiers (Public/Private)
  - Method type (Sub/Function/Property Get/Set/Let)
  - Parameter lists with types
  - Return types for functions
- **Benefit**: Better IntelliSense, go-to-definition, outline view

### 8. Enhanced Argument Handling
- **Current**: Not handled
- **VBA Enhancement**: Detailed parameter parsing with:
  - `Optional` modifier
  - `ByRef`/`ByVal` modifiers
  - `ParamArray` for variable arguments
  - Parameter types and default values
  - Named arguments (`:=` syntax)
- **Benefit**: Better parameter highlighting in signatures and calls

### 9. Type Declarations (Enum/Type)
- **Current**: Only keywords recognized
- **VBA Enhancement**: Complete structures with:
  - Enum member declarations with values
  - Type/UDT member declarations with types
  - Proper scoping for members
- **Benefit**: Better outline view and member highlighting

---

## Priority 4: Control Flow
*Higher complexity*

### 10. Advanced For Loop Parsing
- **Current**: Basic keyword matching
- **VBA Enhancement**: Detailed parsing with:
  - **For...Next** with iterator variable, range, and step
  - **For Each...In** with collection iteration
  - Next with optional variable name
- **Benefit**: Highlights loop components correctly

### 11. Enhanced Select Case
- **Current**: Basic keyword matching
- **VBA Enhancement**: 
  - `Case` expressions with proper scoping
  - `Case Else` handling
  - Expression lists in case conditions
- **Benefit**: Better case structure highlighting

### 12. Open Statement Parsing
- **Current**: Just `Open` keyword
- **VBA Enhancement**: Complete statement with:
  - File path, mode (Input/Output/Append/Binary/Random)
  - Access (Read/Write), Lock, Shared
  - File number (#)
  - Record length (Len)
- **Benefit**: Much better file I/O code readability

---

## Priority 5: Advanced Features
*Complex but valuable*

### 13. Property Chain and Member Access
- **Current**: Not handled
- **VBA Enhancement**: Recognizes object.property chains with:
  - Dot accessor (`.`)
  - Bang accessor (`!`) for collection access
  - Nested chains
  - `Me` keyword
- **Benefit**: Semantic highlighting of object members

### 14. Function vs Sub Call Distinction
- **Current**: All lumped together
- **VBA Enhancement**: Separate patterns for:
  - **Function calls**: with parentheses `func(args)`
  - **Sub calls**: with/without arguments
  - **Sub calls without args**: just the name
- **Benefit**: Different styling for different call types

### 15. Variable and Assignment Tracking
- **Current**: Not handled
- **VBA Enhancement**: 
  - Variable references
  - `Let`/`Set` assignments with proper scoping
  - Assignment operator highlighting
- **Benefit**: Distinguishes declarations from usage

### 16. Labels
- **Current**: Not handled explicitly
- **VBA Enhancement**: Recognizes line labels (`Label:` or `100:`)
- **Benefit**: Better GoTo target visibility

---

## Priority 6: Module Structure
*Specialized*

### 17. Module Headers and Attributes
- **Current**: Generic metadata block
- **VBA Enhancement**: Specific patterns for:
  - `VERSION` declarations
  - `Attribute VB_Name` etc.
  - `Option` statements with specific values
- **Benefit**: Better .cls/.frm file structure

### 18. VB Built-in Constants
- **Current**: Only True/False/Nothing/Null/Empty
- **VBA Enhancement**: Recognizes VB enum constants (vbRed, vbCrLf, etc.)
- **Benefit**: Highlights standard library constants

---

## Priority 7: Polish
*Nice-to-have*

### 19. Expression Parsing
- **Current**: Not structured
- **VBA Enhancement**: Proper expression boundaries with recursive patterns
- **Benefit**: More accurate nested expression highlighting

### 20. Comment Improvements
- **Current**: Basic apostrophe and Rem
- **VBA Enhancement**: 
  - Handles line continuation in comments
  - Better multi-line comment handling
- **Benefit**: More robust comment parsing

---

## Recommended Implementation Order

### Phase 1 (Quick wins)
Items 1-4: String handling, number suffixes, line continuation, line separator

### Phase 2 (Operators)
Items 5-6: Operator categorization and range operator

### Phase 3 (Structure)
Items 7-9: Method signatures, arguments, type declarations

### Phase 4 (Control flow)
Items 10-12: For loops, Select Case, Open statements

### Phase 5 (Advanced)
Items 13-16: Property chains, call distinction, variable tracking, labels

### Phase 6 (Specialized)
Items 17-18: Module headers and VB constants

### Phase 7 (Polish)
Items 19-20: Expression parsing and comment improvements

---

## Implementation Notes

Each phase builds on the previous one and can be tested independently. The phases are designed to:
- Minimize breaking changes
- Provide incremental value
- Allow for testing and validation at each step
- Maintain backward compatibility with existing code

## References

- Source VBA Grammar: `syntaxes/vba.tmLanguage.yaml`
- Target VB6 Grammar: `syntaxes/vb6.yaml-tmlanguage`
