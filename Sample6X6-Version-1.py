import os
import win32com.client as win32
import sys
import requests
def ensure_docm_format(file_path):
    if file_path.endswith('.docm'):
        return file_path  # File is already a .docm, no need to convert
    else:
        # Convert .docx to .docm
        new_file_path = file_path.replace('.docx', '.docm')
        word = win32.DispatchEx("Word.Application")
        doc = word.Documents.Open(file_path)
        doc.SaveAs2(new_file_path, FileFormat=13)  # 13 corresponds to .docm file format
        doc.Close()
        word.Quit()
        return new_file_path
def collect_replacements_chronologically():
    replacements = []
    global market_name_new
    market_name_new = input("Enter the market name to replace 'Xyz': ")
    replacements.append(("Xyz", market_name_new))

    while True:
        try:
            num_companies = int(input("How many companies are there? "))
            break  # Exit the loop if the input is a valid integer
        except ValueError:
            print("Please enter a valid integer for the number of companies.")

    for i in range(1, num_companies + 1):
        company_new = input(f"Enter the name for Company {i:02d}: ")
        replacements.append((f"COMPANY {i:02d}", company_new))

    while True:
        try:
            num_segments = int(input("How many segments are there? "))
            break  # Exit the loop if the input is a valid integer
        except ValueError:
            print("Please enter a valid integer for the number of segments.")

    for i in range(1, num_segments + 1):
        segment_new = input(f"Enter the new name for Segment {i}: ")
        replacements.append((f"Segment {i}", segment_new))

        while True:
            try:
                num_subsegments = int(input(f"How many subsegments are there for {segment_new}? "))
                break  # Exit the loop if the input is a valid integer
            except ValueError:
                print("Please enter a valid integer for the number of subsegments.")

        for j in range(1, num_subsegments + 1):
            subsegment_new = input(f"Enter the new name for Sub-segment {j} of {segment_new}: ")
            replacements.append((f"Seg{i}_Sub{j}", subsegment_new))
            
    print("Let's Begin The Hustle!!")
    return replacements

def add_and_run_macro(file_path, replacements):
    word = win32.DispatchEx("Word.Application")
    word.Visible = True
    doc = word.Documents.Open(file_path)

    replacements_lines = "\n".join([f"replacements({i+1}, 1) = \"{pair[0]}\": replacements({i+1}, 2) = \"{pair[1]}\"" for i, pair in enumerate(replacements)])
    macro_code = f"""
    Sub ReplaceTextInDocumentAndChartData()
        Dim replacements(1 To {len(replacements)}, 1 To 2) As String
        {replacements_lines}
        
        Dim rng As Range
        Dim shp As Shape
        Dim inShp As InlineShape
        Dim hdr As HeaderFooter
        Dim sec As Section
        Dim xlsApp As Object
        Dim xlsSheet As Object
        Dim rngData As Object
        Dim cell As Object

        ' Replace in the main document body
        Set rng = ActiveDocument.Content
        For i = 1 To UBound(replacements, 1)
            If InStr(1, replacements(i, 1), "COMPANY", vbTextCompare) > 0 Then
                rng.Find.Execute FindText:=replacements(i, 1), ReplaceWith:=replacements(i, 2), MatchCase:=True, Replace:=wdReplaceAll
            Else
                rng.Find.Execute FindText:=replacements(i, 1), ReplaceWith:=replacements(i, 2), Replace:=wdReplaceAll
            End If
        Next i

        ' Replace in headers
        For Each sec In ActiveDocument.Sections
            For Each hdr In sec.Headers
                If Not hdr.Exists Then Exit For
                For i = 1 To UBound(replacements, 1)
                    If InStr(1, replacements(i, 1), "COMPANY", vbTextCompare) > 0 Then
                        hdr.Range.Find.Execute FindText:=replacements(i, 1), ReplaceWith:=replacements(i, 2), MatchCase:=True, Replace:=wdReplaceAll
                    Else
                        hdr.Range.Find.Execute FindText:=replacements(i, 1), ReplaceWith:=replacements(i, 2), Replace:=wdReplaceAll
                    End If
                Next i
            Next hdr
        Next sec

        ' Replace in shapes like text boxes
        For Each shp In ActiveDocument.Shapes
            If shp.Type = msoTextBox Then
                For i = 1 To UBound(replacements, 1)
                    If InStr(1, replacements(i, 1), "COMPANY", vbTextCompare) > 0 Then
                        shp.TextFrame.TextRange.Find.Execute FindText:=replacements(i, 1), ReplaceWith:=replacements(i, 2), MatchCase:=True, Replace:=wdReplaceAll
                    Else
                        shp.TextFrame.TextRange.Find.Execute FindText:=replacements(i, 1), ReplaceWith:=replacements(i, 2), Replace:=wdReplaceAll
                    End If
                Next i
            End If
        Next shp

        ' Replace data in charts embedded as InlineShapes
        For Each inShp In ActiveDocument.InlineShapes
            If inShp.Type = wdInlineShapeChart Then
                inShp.Chart.ChartData.Activate ' Activate the chart data
                Set xlsApp = inShp.Chart.ChartData.Workbook.Application
                Set xlsSheet = xlsApp.Worksheets(1)
                Set rngData = xlsSheet.UsedRange
                For Each cell In rngData
                    For i = 1 To UBound(replacements, 1)
                        If CStr(cell.Value) = replacements(i, 1) Then
                            cell.Value = replacements(i, 2)
                        End If
                    Next i
                Next cell
                inShp.Chart.ChartData.Workbook.Close SaveChanges:=True
            End If
        Next inShp

        ' Replace data in charts embedded as Shapes
        For Each shp In ActiveDocument.Shapes
            If shp.HasChart Then
                shp.Chart.ChartData.Activate ' Activate the chart data
                Set xlsApp = shp.Chart.ChartData.Workbook.Application
                Set xlsSheet = xlsApp.Worksheets(1)
                Set rngData = xlsSheet.UsedRange
                For Each cell In rngData
                    For i = 1 To UBound(replacements, 1)
                        If CStr(cell.Value) = replacements(i, 1) Then
                            cell.Value = replacements(i, 2)
                        End If
                    Next i
                Next cell
                shp.Chart.ChartData.Workbook.Close SaveChanges:=True
            End If
        Next shp

        For Each sec In ActiveDocument.Sections
            For Each hdr In sec.Headers
                If hdr.Exists Then
                    ' Replace text in the header's range
                    For i = 1 To UBound(replacements, 1)
                        If InStr(1, replacements(i, 1), "COMPANY", vbTextCompare) > 0 Then
                            hdr.Range.Find.Execute FindText:=replacements(i, 1), ReplaceWith:=replacements(i, 2), MatchCase:=True, Replace:=wdReplaceAll
                        Else
                            hdr.Range.Find.Execute FindText:=replacements(i, 1), ReplaceWith:=replacements(i, 2), Replace:=wdReplaceAll
                        End If
                    Next i
                            
                    ' Additional loop to handle textboxes within the header
                    For Each shp In hdr.Shapes
                        If shp.Type = msoTextBox Then
                            For i = 1 To UBound(replacements, 1)
                                If shp.TextFrame.HasText Then
                                    If InStr(1, replacements(i, 1), "COMPANY", vbTextCompare) > 0 Then
                                        shp.TextFrame.TextRange.Find.Execute FindText:=replacements(i, 1), ReplaceWith:=replacements(i, 2), MatchCase:=True, Replace:=wdReplaceAll
                                    Else
                                        shp.TextFrame.TextRange.Find.Execute FindText:=replacements(i, 1), ReplaceWith:=replacements(i, 2), Replace:=wdReplaceAll
                                    End If
                                End If
                            Next i
                        End If
                    Next shp
                End If
            Next hdr
        Next sec

        ' Update the Table of Contents
        If ActiveDocument.TablesOfContents.Count > 0 Then
            ActiveDocument.TablesOfContents(1).Update
        End If
    End Sub
    """

    try:
        vbaproject = doc.VBProject
        vbamodule = vbaproject.VBComponents.Add(1)
        vbamodule.CodeModule.AddFromString(macro_code)
        word.Application.Run("ReplaceTextInDocumentAndChartData")
        
        # Determine the executable's directory for the output
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))

        new_file_name = f"Sample_{market_name_new}.docx"
        edited_file_path = os.path.join(application_path, new_file_name)
        doc.SaveAs2(FileName=edited_file_path, FileFormat=16)
        print(f"File saved to {edited_file_path}")
    except Exception as e:
        print(f"Error: {e}")
    finally:
        doc.Close(False)
        word.Quit()
    input("Press Enter to exit...")

def fetch_license_key(github_raw_url):
    """Fetch the license key from a JSON file hosted on GitHub."""
    try:
        response = requests.get(github_raw_url)
        response.raise_for_status()  # This will raise an exception for HTTP errors
        data = response.json()
        return data.get("license_key", "")
    except Exception as e:
        print(f"Failed to fetch license key: {e}")
        return ""

def validate_license(expected_key, github_raw_url):
    """Validate the license key fetched from GitHub against the expected key."""
    actual_key = fetch_license_key(github_raw_url)
    return actual_key == expected_key

# Test the functions

if __name__ == "__main__":
    GITHUB_RAW_URL = "https://raw.githubusercontent.com/BhuwanAwasthi/Orb/main/sample_key.json"
    EXPECTED_LICENSE_KEY = "Trycopymeandfeelmywrath"  # Adjusted to the correct expected key
    attempt_count = 0
    license_valid = False
    
    print("Checking license...")
    while attempt_count < 3 and not license_valid:
        if validate_license(EXPECTED_LICENSE_KEY, GITHUB_RAW_URL):
            print("License validation successful. Proceeding with the program...")
            license_valid = True
        else:
            attempt_count += 1
            if attempt_count < 3:
                print("Authorization failed. Click here to try again.")
            else:
                print("License validation failed after 3 attempts. Please contact Bhuwan for this issue ASAP.")
                break  # Exit the loop if license is not valid after 3 attempts
    
    if not license_valid:
        # Exit the script if license validation fails
        sys.exit()

    current_dir = os.path.dirname(os.path.abspath(__file__))
    print("Hi There!!")
    print("© 2023 Adroit Market Research. All Rights Reserved.")
    print("Note: This software solely belongs to Adroit Market Research \n If you are not a part of the team, Close this window immediately!!")
    choice = input("Which sample report are you working on? \n Enter 1 for Orbis and 2 for Adroit: ")
    if choice == "1":
        original_file_path = os.path.join(current_dir, "Orbis_sample.docm")
    elif choice == "2":
        original_file_path = os.path.join(current_dir, "Adroit_sample.docm")
    else:
        print("Invalid choice. Please enter either '1' for 'Orbis' or '2' for 'Adroit'.")
        exit()
    
    file_path = ensure_docm_format(original_file_path)
    replacements = collect_replacements_chronologically()
    add_and_run_macro(file_path, replacements)

