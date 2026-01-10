1. DOS batch script is case insensitive.  
   All command and parameter can be upper case, lower case, or hybrid without affecting the result.  

2. A batch script is usually with `.bat` or `.cmd` extension name.  

3. `echo` controls whether include input lines in the output.  
   `echo off` to turn off the echo for all lines below it, until encouter `echo on`.  
   Starts a new line with `@` to turn off echo for that single line.  

4. `rem` or `::` leads a line of comment.  

5. `help <command>` provides instruction for the <command>.  E.g., `help dir`.  

6. Use `set` to assign varibles in a batch script.  
   E.g., `set var=abc` will assign `"abc"` to `%var%`.  
   Don't adding extra space before and after `=`, otherwise it considered part of varable name or value.  
   Use `%` to surround the varible name when using it, e.g., `echo %var%`.  
   `help set` would provide more detail.  

7. `if` is for branching control.  
   Use `==` to compare values and evaluate condition.  
   `help if` for more detail.  

8. `for` is for loop control.  
   `help for` for more detail.  

9. `%~dp0` is to obtain drive letter and path of the command (%0, first token).  
   Token is each string in command-line separated by space (by default).  
   The command itself is usually %0, first argument is %1, and so forth.  
   To expand a token representing a file path:
   - `d` means drive letter
   - `p` means path only (no drive letter or file name)
   - `f` means full path
   - `n` means file name
   - `x` means extension
   
   `help for` includes more detail on this, and usage of other expansion of tokens.  

11. To drag and drop multiple files into a batch script and process them, using a loop.  
    Example:  
    ```
    :loop
    if [%1]==[] goto :endLoop
    echo %1
    shift
    goto loop
    :endLoop
    ```  

12. 
