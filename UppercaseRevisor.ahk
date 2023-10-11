#InstallKeybdHook
SetKeyDelay, -1 ;Makes the edit instantaneous. Doesn't seem to work though. Anyway, it's fast enough
SetBatchLines, -1

;CONSTANTS:
global bufferSize := 40 ;The number characters plus shift key events to keep track of
global mymargin := 300 ;The minimum number of milliseconds between two keystrokes so that they can be considered definitely in the right order
global shiftPressed := "{Shift Down}"
global shiftReleased := "{Shift Up}"
global uncapitalizedCharacters := "abcdefghijklmnopqrstuvwxyz1234567890-=[]\;',./"    ;The characters eligible for substitution for their respective capitalized characters
global capitalizedCharacters := "ABCDEFGHIJKLMNOPQRSTUVWXYZ!@#$%^&*()_+{}|:""<>?" ;is one character longer due to double quotes being represented as "" rather than "
global newline := "`r"

;variables:
global keyBuffer := Object()
global timestampBuffer := Object()
global index := 0
global successiveInvocationCount := 1
global lastExecutedToggleToken := -1

;hotkeys with noticable effect:
^q::Invoke()
; F12::OutputBuffers()

;outputs a message box with debugging info on the buffers
OutputBuffers()
{
    global
    i := 0
    str := ""
    while (i < index)
    {
        str := % str . keyBuffer[i] . " at " . timestampBuffer[i] . newline
        i++
    }
    MsgBox %str%
}

;main function, see the https://github.com/JeroenBos/UppercaseRevisor/wiki for details
Invoke()
{
    global bufferSize
    global successiveInvocationCount

    pressIndex := -1
    releaseIndex := -1

    if(not FindShiftPressAndRelease(pressIndex, releaseIndex))
    {
        return
    }

    if(releaseIndex - pressIndex = 1)
    {
        x := false
        if(lastExecutedToggleToken = 4)
        {
            ;undoes the capitalization by a previous consecutive invocation
            Decapitalize(Mod(releaseIndex + 1, bufferSize))
            x := false
        }
        else if(lastExecutedToggleToken = 5)
        {
            ;undoes the capitalization by a previous consecutive invocation
            Decapitalize(Mod(pressIndex - 1 + bufferSize, bufferSize))
            x := true
        }
        else if(IsWithinMarginOf(releaseIndex))
        { 
            x := true
        }
        else if(IsWithinMarginOf(Mod(pressIndex - 1 + bufferSize, bufferSize)))
        {
            x := false
        }
        else if(Mod(releaseIndex + 1, bufferSize) <> index) ;checks existence of key after shift released
        {
            if(Mod(pressIndex - 1 + bufferSize, bufferSize) <> index) ;checks existence of key before shift pressed
            {
                ;modifies whichever was closest to a shift key press or release
                x := timestampBuffer[Mod(releaseIndex + 1, bufferSize)] - timestampBuffer[releaseIndex] < timestampBuffer[pressIndex] - timestampBuffer[Mod(pressIndex - 1 + bufferSize, bufferSize)]
            }
            else
            {
                x := true
            }
        }
        else if(Mod(releaseIndex - 1 + bufferSize, bufferSize) <> index) ;checks existence of key before shift pressed
        {
            x := false
        }
        else
        {
            ; do nothing, there are no keys other than the shift keys
            return
        }
        ;capitalizes either the character before or after, i.e. x, the shift press and release
        NoCharacterExecute(x, pressIndex, releaseIndex)
    }
    else if(releaseIndex - pressIndex = 2)
    {
        if(lastExecutedToggleToken <> -1)
        {
            UndoLastToggleSingleCharacterWithinBounds(Mod(pressIndex + 1, bufferSize))
        }
        ToggleSingleCharacterWithinBounds(Mod(pressIndex + 1, bufferSize))
    }
    else if(pressIndex < Mod(releaseIndex - 1 - successiveInvocationCount + bufferSize, bufferSize))
    {
        Decapitalize(Mod(releaseIndex - 1 - successiveInvocationCount + bufferSize, bufferSize))
    }
    else if(pressIndex = Mod(releaseIndex - 1 - successiveInvocationCount + bufferSize, bufferSize) ;whether all successive capitalized letters have been decapitalized
        && Mod(releaseIndex + 1, bufferIndex) <> index) ;and the shift release wasn't the last action
    {
        numberOfCharactersToCapitalize := releaseIndex - pressIndex - 1
        i := pressIndex + 1 

        while(numberOfCharactersToCapitalize >= 0)
        {
            Capitalize(i)
            numberOfCharactersToCapitalize--
            i++
        }
        Capitalize(Mod(releaseIndex + 1, bufferSize))
    }
    else
    {
        if(Mod(pressIndex + 2 + successiveInvocationCount, bufferIndex) <> index) ;whether there is another succesiveInvocationCount'th character after the shift release
        {
            Capitalize(Mod(pressIndex + 2 + successiveInvocationCount, bufferSize)) 
        }
    }

    successiveInvocationCount++
    return
}

NoCharacterExecute(x, pressIndex, releaseIndex)
{
    if(x)
    {
        Capitalize(Mod(releaseIndex + 1, bufferSize))
        lastExecutedToggleToken := 4
    }
    else
    {
        Capitalize(Mod(pressIndex - 1 + bufferSize, bufferSize))
        lastExecutedToggleToken := 5
    }
}

AppendToBuffer(key)
{
    global
    keyBuffer[index] := key
    timestampBuffer[index] := A_TickCount
    index := Mod(index + 1, bufferSize) 
    if (key <> shiftPressed && key <> shiftReleased)
    {
        successiveInvocationCount := 0
    }
    lastExecutedToggleToken := -1
    return
}

;removes the last character input from the buffer. Does not remove buffered shift key presses or releases
PopBuffer()
{
    global

    lastExecutedToggleToken := -1

    characterToRemoveIndex := FindLastNonShiftBufferedAction()

    if(characterToRemoveIndex = -1)
    {
        return ; only shift key actions are buffered
    }
    else
    {
        index--
        BuffersRemoveAt(characterToRemoveIndex)
    }
}

FindLastNonShiftBufferedAction()
{
    unboundedIndex := index + bufferSize - 1 ;
    while index <= unboundedIndex
    {
        boundedIndex := Mod(unboundedIndex, bufferSize)

        if(keyBuffer[boundedIndex] <> shiftPressed && keyBuffer[boundedIndex] <> shiftReleased)
        {
            return boundedIndex
        }

        unboundedIndex--
    }

    return -1
}

;removes the element at the specified index in the buffer (and in the time tick buffer)
BuffersRemoveAt(indexToRemove)
{
    global keyBuffer 
    global timestampBuffer

    if(indexToRemove < index)
    {
        numberOfElementsToMove := index - indexToRemove
    }
    else
    {
        numberOfElementsToMove := bufferSize - index + indexToRemove - 1
    }

    unboundedIndex := indexToRemove
    while unboundedIndex <> indexToRemove + numberOfElementsToMove
    {
        boundedIndex := Mod(unboundedIndex, bufferSize)
        boundedNextIndex := Mod(unboundedIndex + 1, bufferSize)

        keyBuffer[boundedIndex] := keyBuffer[boundedNextIndex]
        timestampBuffer[boundedIndex] := timestampBuffer[boundedNextIndex]

        unboundedIndex++
    }
}

ClearBuffers()
{
    keyBuffer := Object()
    timestampBuffer := Object()
    index := 0
    successiveInvocationCount := 0
    lastExecutedToggleToken := -1
}

RemoveCharacters(start, length)
{
    global
    i := start - 1
    while length <> 0
    {
        bufferedKey := keyBuffer[i]
        if(not bufferedKey = shiftPressed && not bufferedKey = shiftReleased)
        { 
            Send {Backspace} 
        }

        i := Mod(i - 1 + bufferSize, bufferSize)
        length--
    }
    return
}

InsertCharacter(character)
{
    Send %character%
}
InsertBufferedCharacter(i)
{
    InsertCharacter(keyBuffer[i]) 
}
InsertCharacters(i, count)
{
    if(count = 0)
    {
        return
    } 
    else
    {
        InsertBufferedCharacter(i)
        InsertCharacters(Mod(i + 1, bufferSize), count - 1)
    }
}

Capitalize(i) ;i is an index in the buffers
{
    global keyBuffer
    characterToCapitalize := keyBuffer[i]
    substitute := GetCapitalizedForm(keyBuffer[i])
    Replace(i, substitute)
}
Decapitalize(i)
{
    global keyBuffer
    Replace(i, GetDecapitalizedForm(keyBuffer[i]))
}
Replace(i, substitute) ;i is an index in the buffers, substitute is the character to insert there
{
    global
    if(substitute = "")
    {
        return
    }
    debug_mostRecentKey := keyBuffer[index - 1]
    toRemoveCount := Mod(index - i + bufferSize, bufferSize) ;the number of characters (including shift presses) to remove to just include removing the character at bufferindex i
    if(toRemoveCount = 0)
    {
        toRemoveCount := bufferSize
    }

    RemoveCharacters(index, toRemoveCount)
    InsertCharacter(substitute)
    substituteIndex := Mod(index - toRemoveCount + bufferSize, bufferSize)
    currentlyAtSubsituteIndex := keyBuffer[substituteIndex]
    keyBuffer[substituteIndex] := substitute ;also changes the character in the buffer, to ensure that any successive modification takes into account the substitute
    InsertCharacters(Mod(i + 1, bufferSize), toRemoveCount - 1) 
}

FindShiftPressAndRelease(ByRef pressIndex, ByRef releaseIndex)
{
    global ; index bufferSize actionBuffer keyBuffer
    unboundedIndex := index + bufferSize - 1 ;-1 to start at the last character (since i points to the index where the next character would be placed)
    releaseFound := false
    pressFound := false
    while index <= unboundedIndex
    {
        boundedIndex := Mod(unboundedIndex, bufferSize)
        key := keyBuffer[boundedIndex]
        debug_timestamp := timestampBuffer[boundedIndex]
        if(key = shiftPressed)
        {
            if(pressFound)
            {
                ;two consecutive shift pressed found, as it, in is still being pressed
            }
            else if(releaseFound)
            {
                ;found the correct sequence: now we found a shift press for which we already have found a shift release.
                pressIndex := boundedIndex
                return true
            }
            else
            {
                ;we found the latest shift press
                pressFound := true 
                pressIndex := boundedIndex
            }
        }
        else if(key = shiftReleased)
        {
            if(releaseFound)
            {
                return false ;two consecutive shift releases found
            }
            else if(pressFound)
            {
                ;found that during invocation of this programme, shift is down. We'll ignore that shift press and look for a next shift press before the shift release found now       
                ;pressFound = false; EDIT: actually, the implementation isn't adapted to this scenario. Just fail:
                return false
            }
            else
            {
                ;found a shift release, for which we still need to search the corresponding press
                releaseFound := true
                releaseIndex := boundedIndex
            } 
        } 
        else
        {
            ;the key represents some other character
        } 
        unboundedIndex--
    }
    ;exhausted all buffered keys, to no avail
    return false
}

GetCapitalizedForm(character)
{
    global
    return GetOtherForm(character, uncapitalizedCharacters, capitalizedCharacters)
}
GetDecapitalizedForm(character)
{
    global
    return GetOtherForm(character, capitalizedCharacters, uncapitalizedCharacters)
}
GetOtherForm(character, keys, values)
{
    i := IndexOf(keys, character)
    if(i <> -1)
    {
        return SubStr(values, i, 1)
    }

    ;the character wasn't there... hmmm... ;actually, intended for spacebar, tab, enter
    ;return the original character, so that in the end no effect is produced, since it is replaced by itself 
    return character
}
At(string, index) ;index is one-based, like in any non-self-respecting programming language
{
    return SubStr(string, index, 1)
}

IndexOf(string, char)
{
    i := 1
    length := StrLen(string)
    while i <= length
    {
        if(At(string, i) = char)
        {
            return i
        }
        i++
    }
    return -1 
}

;gets whether the key stroke at the specified index is within the allowed temporal margin from the next key stroke, if any
IsWithinMarginOf(i) ;i is the index of the first of the two 
{
    global mymargin
    t1 := timestampBuffer[i]
    t2 := timestampBuffer[Mod(i + 1, bufferSize)]
    if(not t1 OR not t2)
        return false ;then the index of either one is out of range, in which case there is no margin to consider, hence "false"
    return mymargin >= t2 - t1 
}

ToggleSingleCharacterWithinBounds(i)
{
    global
    if(lastExecutedToggleToken <= 0) ;can include -1
    {
        if(IsWithinMarginOf(Mod(i + 1, bufferSize)))
        {
            Capitalize(Mod(i + 2, bufferSize))
            lastExecutedToggleToken := 1
            return 
        }
    }
    if(lastExecutedToggleToken <= 1)
    {
        if(IsWithinMarginOf(Mod(i - 2 + bufferSize, bufferSize)))
        {
            Capitalize(Mod(i - 2 + bufferSize, bufferSize))
            lastExecutedToggleToken := 2
            return 
        }
    }
    if(lastExecutedToggleToken <> 0)
    {
        Decapitalize(i)
        lastExecutedToggleToken := 0
    }
    return
}

UndoLastToggleSingleCharacterWithinBounds(i) ;i is the character of the single character between the shift press and release
{
    global
    if(lastExecutedToggleToken = 0)
    {
        ;undo decapitalization of the only character between shifts
        Capitalize(i)
    }
    else if(lastExecutedToggleToken = 1)
    {
        ;undo capitalization of the character after the shift release. Checking margin is redudant, otherwise the lastExecutedToggleToken would have been something different
        Decapitalize(Mod(i + 2, bufferSize))
    }
    else if(lastExecutedToggleToken = 2)
    {
        ;undo capitalization of the character before the shift press. Checking margin is redudant, otherwise the lastExecutedToggleToken would have been something different
        Decapitalize(Mod(i - 2 + bufferSize, bufferSize))
    }
}

;key strokes that alter the behavior completely without fix:
~*Home:: ; '*' means irrespective of modifier keys
~*End::
~*PGDN::
~*PGUP::
~*UP::
~*Down::
~*Right::
    ~*Backspace::ClearBuffers()

    ;key strokes that have special handling
~Space:: 
    ~+Space::AppendToBuffer(" ")
    ~Tab::AppendToBuffer("{Tab}")
    ~Enter::AppendToBuffer("{Enter}")
~Left::
    ~+Left::index--
~BackSpace::
    ~+Backspace::PopBuffer()

    ;SHIFT: huh: apparently the tilde here matters: ~Shift::MsgBox displays a messagebox on shift down, but Shift::MsgBox displays it on shift up :/ Anyway, with tilde is the desired behavior so just .... yeah... whatever. it works
~Shift::
    {
        if(keyBuffer[index] <> shiftPressed) ;prevents adding successive shift pressed (in case the shift key is being held down). This requirement is optional, but the buffer could quickly overflow with just shift presses
        {
            AppendToBuffer(shiftPressed)
        }
        return
    }
    ~Shift Up::AppendToBuffer(shiftReleased)

    ;buffer all typed characters
~a::
~b::
~c::
~d::
~e::
~f::
~g::
~h::
~i::
~j::
~k::
~l::
~m::
~n::
~o::
~p::
~q::
~r::
~s::
~t::
~u::
~v::
~w::
~x::
~y::
~z::
~`::
~1::
~2::
~3::
~4::
~5::
~6::
~7::
~8::
~9::
~0::
~-::
~=::
~[::
~]::
~\::
    ~;::
~'::
~,::
~.::
    ~/::AppendToBuffer(Substr(A_ThisHotkey, 2))

~+A::
~+B::
~+C::
~+D::
~+E::
~+F::
~+G::
~+H::
~+I::
~+J::
~+K::
~+L::
~+M::
~+N::
~+O::
~+P::
~+Q::
~+R::
~+S::
~+T::
~+U::
~+V::
~+W::
~+X::
~+Y::
    ~+Z:: AppendToBuffer(Substr(A_ThisHotkey, 3))

    ~+1:: AppendToBuffer("!")
    ~+2:: AppendToBuffer("@")
    ~+3:: AppendToBuffer("#")
    ~+4:: AppendToBuffer("$")
    ~+5:: AppendToBuffer("%")
    ~+6:: AppendToBuffer("^")
    ~+7:: AppendToBuffer("&")
    ~+8:: AppendToBuffer("*")
    ~+9:: AppendToBuffer("(")
    ~+0:: AppendToBuffer(")")
    ~+-:: AppendToBuffer("_")
    ~+=:: AppendToBuffer("+")
    ~+[:: AppendToBuffer("{")
    ~+]:: AppendToBuffer("}")
    ~+\:: AppendToBuffer("|")
    ~+;:: AppendToBuffer(":")
    ~+':: AppendToBuffer("""")
    ~+,:: AppendToBuffer("<")
    ~+.:: AppendToBuffer(">")
    ~+/:: AppendToBuffer("?")
