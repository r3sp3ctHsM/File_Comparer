def compare_texts(text1, text2):
    """Compare two texts and return the differences."""
    differences = []
    text1_lines = text1.splitlines()
    text2_lines = text2.splitlines()

    # Compare line by line
    for i, (line1, line2) in enumerate(zip(text1_lines, text2_lines)):
        if line1 != line2:
            differences.append(f"Line {i+1}: '{line1}' != '{line2}'")

    # Check if one text has extra lines
    if len(text1_lines) > len(text2_lines):
        for i in range(len(text2_lines), len(text1_lines)):
            differences.append(f"Line {i+1}: '{text1_lines[i]}' (extra in text1)")
    elif len(text2_lines) > len(text1_lines):
        for i in range(len(text1_lines), len(text2_lines)):
            differences.append(f"Line {i+1}: '{text2_lines[i]}' (extra in text2)")

    return differences

def compare_fonts(fonts1, fonts2):
    """Compare font styles between two sets of fonts."""
    font_differences = []

    # Compare font lists
    for i, (font1, font2) in enumerate(zip(fonts1, fonts2)):
        if font1 != font2:
            font_differences.append(f"Font difference at position {i+1}: '{font1}' != '{font2}'")

    # Check for extra fonts in either list
    if len(fonts1) > len(fonts2):
        for i in range(len(fonts2), len(fonts1)):
            font_differences.append(f"Extra font in fonts1 at position {i+1}: '{fonts1[i]}'")
    elif len(fonts2) > len(fonts1):
        for i in range(len(fonts1), len(fonts1)):
            font_differences.append(f"Extra font in fonts2 at position {i+1}: '{fonts2[i]}'")

    return font_differences