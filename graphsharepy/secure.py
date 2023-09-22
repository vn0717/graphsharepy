def wipe_mem(value):
    """
    Function to wipe a string variable in memory.  It replaces
    variable content with zeros to secure the information the 
    variable use to hold.

    Args:
        value (STRING): The string variable to wipe the data from

    Returns:
        STRING: A string of zeros
    """
    length = len(value)
    value = "0"
    for i in range(length-1):
        value += "0"
    return value



def wipe_subval(key, string):
    """
    Function to wipe the memory of a string that represents a dictonary.
    It finds the where the key is defined in the string and then wipes the values after.

    Args:
        key (STRING): The dictonary key that holds the information that needs to be secured
        string (STRING): The string that represents a dictonary

    Returns:
        STRING : The string that represents a dictonary with the memory secured
    """
    loc = string.find(key)
    if loc != -1:
        length = len(key)+1
        loc += length
        sub_string = string[loc:]
        end = sub_string.find(",")
        word = sub_string[:end]
        length_word = len(word)
        for i in range(length_word):
            word = word[:i] + "0" + word[i + 1:]
            string = string[:i+loc] + "0" + string[i+loc + 1:]

    return string