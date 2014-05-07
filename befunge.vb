Option Explicit
Option Base 0

Enum Directions
    None = 0
    Up
    Down
    Left
    Right
End Enum
       
Sub befunge2()
    Dim end_program As Boolean, text_input_mode As Boolean
    Dim origin As Range, program_counter As Range
    Dim execution_direction As Directions
    Dim stack As Stacks
    
    Set origin = Range("B2")
    Set program_counter = origin
    execution_direction = Right
    end_program = False
    text_input_mode = False
    
    Set stack = New Stacks
    stack.init
    
    Do While Not end_program
        delay 0.1
        program_counter.Activate
        parse_cell CStr(program_counter.value), program_counter, origin, _
                        execution_direction, stack, end_program, text_input_mode
    Loop
    
    Debug.Print "Program terminated."
End Sub

Sub parse_cell(contents As String, ByRef program_counter As Range, origin As Range, ByRef execution_direction As Directions, _
               ByRef stack As Stacks, ByRef end_program As Boolean, ByRef text_input_mode As Boolean)
    
    If text_input_mode Then
        'if another " is found set text_input_mode to false and move to next cell
        If (program_counter.value = Chr(34)) Then
            text_input_mode = False
            update_pc program_counter, execution_direction
        'otherwise push the ascii code for the symbol in the cell and then move to next cell
        Else
            stack.push Asc(program_counter.value)
            update_pc program_counter, execution_direction
        End If
    Else
        'These variables are for get and put calls
        Dim x As Integer, y As Integer
        Dim v As String
        
        Select Case contents
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" 'Push this number on the stack
                stack.push CInt(contents)
                update_pc program_counter, execution_direction
            Case "+"       'Addition: Pop a and b, then push a+b
                stack.add
                update_pc program_counter, execution_direction
            Case "-"       'Subtraction: Pop a and b, then push b-a
                stack.subtract
                update_pc program_counter, execution_direction
            Case "*"       'Multiplication: Pop a and b, then push a*b
                stack.multiply
                update_pc program_counter, execution_direction
            Case "/"       'Integer division: Pop a and b, then push b/a, rounded down.
                stack.divide
                update_pc program_counter, execution_direction
            Case "%"       'Modulo: Pop a and b, then push the remainder of the integer division of b/a.
                stack.modulo
                update_pc program_counter, execution_direction
            Case "!"       'Logical NOT: Pop a value. If the value is zero, push 1; otherwise, push zero."
                stack.logical_negate
                update_pc program_counter, execution_direction
            Case "`"       'Greater than: Pop a and b, then push 1 if b>a, otherwise zero."
                stack.greaterthan
                update_pc program_counter, execution_direction
            Case ">"       'Start moving right"
                execution_direction = Right
                update_pc program_counter, execution_direction
            Case "<"       'Start moving left"
                execution_direction = Left
                update_pc program_counter, execution_direction
            Case "^"       'Start moving up"
                execution_direction = Up
                update_pc program_counter, execution_direction
            Case "v"       'Start moving down"
                execution_direction = Down
                update_pc program_counter, execution_direction
            Case "?"       'Start moving in a random cardinal direction"
                execution_direction = 1 + CInt((3 * Rnd()))
                update_pc program_counter, execution_direction
            Case "_"       'Pop a value; move right if value=0, left otherwise"
                If stack.pop = 0 Then
                    execution_direction = Right
                Else
                    execution_direction = Left
                End If
                update_pc program_counter, execution_direction
            Case "|"       'Pop a value; move down if value=0, up otherwise"
                If stack.pop = 0 Then
                    execution_direction = Down
                Else
                    execution_direction = Up
                End If
                update_pc program_counter, execution_direction
            Case Chr(34)   '(double quote) Start string mode: push each character's ASCII value all the way up to the next "
                text_input_mode = True
                update_pc program_counter, execution_direction
            Case ":"       'Duplicate value on top of the stack"
                stack.dup
                update_pc program_counter, execution_direction
            Case "\"       'Swap two values on top of the stack"
                stack.swap
                update_pc program_counter, execution_direction
            Case "$"       'Drop value from the stack"
                stack.drop
                update_pc program_counter, execution_direction
            Case "."       'Pop value and output as an integer"
                'Append value to contents of AU3
                Range("AU3").value = CStr(Range("AU3").value) & CStr(stack.pop)
                update_pc program_counter, execution_direction
            Case ","       'Pop value and output as ASCII character"
                'Append value to contents of AU3
                Range("AU3").value = CStr(Range("AU3").value) & CStr(Chr(stack.pop))
                update_pc program_counter, execution_direction
            Case "#"       'Trampoline: Skip next cell"
                update_pc program_counter, execution_direction
                update_pc program_counter, execution_direction
            Case "p"       'A "put" call (a way to store a value for later use). Pop y, x and v, then change the character at the position (x,y) in the program to the character with ASCII value v
                v = Chr(stack.pop)  'Convert ascii code to character
                x = stack.pop + 1   'range.item(1,1) gives you the top left item of range or just range itself if it is a singel cell
                y = stack.pop + 1   'co-ordinates are zero-based to add 1s to algin them
                                
                origin.Item(x, y).value = v
                update_pc program_counter, execution_direction
            Case "g"       'A "get" call (a way to retrieve data in storage). Pop y and x, then push ASCII value of the character at that position in the program
                x = stack.pop + 1   'range.item(1,1) gives you the top left item of range or just range itself if it is a singel cell
                y = stack.pop + 1   'co-ordinates are zero-based to add 1s to algin them
                                
                stack.push (CStr(origin.Item(x, y).value))
                update_pc program_counter, execution_direction
            Case "&"       'Ask user for a number and push it
                update_pc program_counter, execution_direction
            Case "~"       'Ask user for a character and push its ASCII value
                update_pc program_counter, execution_direction
            Case "o"       'Non-standard (from Forth) - copy second item on stack to tos
                stack.over
                update_pc program_counter, execution_direction
            Case ";"       'Output TOS without popping it (non-standard, like .s in Forth)
                'Append value to contents of AU3
                stack.dup
                Range("AU3").value = CStr(Range("AU3").value) & CStr(stack.pop)
                update_pc program_counter, execution_direction
            Case "e"       'Exponentiation - pop a and b then push b^a
                stack.pow
                update_pc program_counter, execution_direction
            Case "@"       'End program
                end_program = True
            Case " ", ""       '(space)/blank No-op. Does nothing
                update_pc program_counter, execution_direction
            Case Else
                Err.Raise vbObjectError + 100, "Befunge", "Unrecognised instruction"
        End Select
    End If
End Sub

Sub update_pc(ByRef program_counter As Range, execution_direction As Directions)
    Select Case execution_direction
        Case Up
            Set program_counter = program_counter.Item(0, 1)
        Case Down
            Set program_counter = program_counter.Item(2, 1)
        Case Left
            Set program_counter = program_counter.Item(1, 0)
        Case Right
            Set program_counter = program_counter.Item(1, 2)
    End Select
End Sub

Sub delay(seconds As Single)
    Dim start As Single
    start = Timer
    Do While Timer - start < seconds
        DoEvents
    Loop
End Sub
