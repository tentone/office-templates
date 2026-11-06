## Unreleased
 - Fix dependency conflicts by aligning the build plugins with Rollup 3 while keeping versions close to the previous releases.
 - Update TypeScript to 5.6 for compatibility with the existing Node.js type definitions.
 - Forward Rollup resolution options when resolving the browser stream polyfill to prevent the build from exiting early.

## 1.0.0
 - First release
 - Forked form "docx-templates" V4.15.0
 - Remove QUERY command
 - Allow for nested IF statement
 - Keep `$idx` available inside nested IF conditions and contents, using the innermost active FOR index and restoring the outer index after an inner FOR ends.
 - Preserve table cells, their formatting, and their rows when a false `IF` block is contained within a cell, in both DOCX and PPTX. `FOR` row and column behavior is unchanged.
 
