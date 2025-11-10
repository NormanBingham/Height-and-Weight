import openpyxl
import os
import argparse
from fillpdf import fillpdfs
from datetime import datetime

def generate_pdfs(excel_file, pdf_5500, pdf_5501, output_dir, custom_date=None):
    current_date = custom_date or datetime.today().strftime('%Y%m%d')

    HEIGHT_WEIGHT_PASS = (
        "The Soldier has met the Army's height and weight standards as outlined in AR 600-9. "
        "The Soldier's weight of {weight} pounds is within the maximum allowable weight of {max_weight} pounds."
        "The Soldier is encouraged to maintain current fitness and body composition levels to support mission readiness."
    )

    MET_STANDARD = (
        "The individual has met the Army's body fat standards as outlined in AR 600-9. "
        "The Soldier's body fat percentage was {body_fat_percentage}%, which is within the standard of {body_fat_standard}%."
        "The Soldier is in compliance with the Army Body Composition Program (ABCP) "
        "and requires no further action. Maintain focus on overall health and physical readiness."
    )

    DID_NOT_MEET_STANDARDS = (
        "The individual has exceeded the allowable body fat standards as outlined in AR 600-9. "
        "The Soldier's body fat percentage was {body_fat_percentage}%, which exceeds the standard of {body_fat_standard}%."
        "The Soldier is noncompliant with the Army Body Composition Program (ABCP) and must be enrolled in the program. "
        "Soldier will adhere to the requirements outlined in AR 600-9 to achieve compliance. "
    )

    ACFT_FAIL_HEIGHT_WEIGHT_PASS = (
        "The Soldier exceeded the Army's height and weight standards outlined in AR 600-9; however, "
        "the Soldier achieved a total score of 540 or above on the Army Combat Fitness Test (ACFT), "
        "with a minimum of 80 points in each event. As such, the Soldier is exempt from being flagged or enrolled in the Army Body Composition Program (ABCP)."
    )

    workbook = openpyxl.load_workbook(excel_file, data_only=True)
    sheet = workbook.active

    custom_mapping = {
        0: "NAME",
        1: "RANK",
        3: "AGE",
        4: "HEIGHT",
        5: "WEIGHT",  
        9: "FIRST",
        10: "SECOND",
        11: "THIRD",
        12: ["AVERAGE", "AVERAGE_1"],  
        13: "BODY FAT PERCENTAGE",
        17: "PREPARED BY",
        18: "PREP_BY_RANK",
        20: "APPROVED BY SUPERVISOR",
        21: "APPR_BY_RANK",
    }

    prepared_by = sheet.cell(row=2, column=1).value
    prep_by_rank = sheet.cell(row=2, column=2).value
    preparer_initials = sheet.cell(row=2, column=3).value
    approved_by_supervisor = sheet.cell(row=4, column=1).value
    appr_by_rank = sheet.cell(row=4, column=2).value

    print("Starting PDF generation...")

    for row in sheet.iter_rows(min_row=6, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
        if not row[0].value:
            print(f"Stopping at row {row[0].row} as the first column is empty.")
            break  

        gender = row[2].value  # Column 3 (index 2) for gender
        if gender == "M":
            pdf_template = pdf_5500
            form_field = fillpdfs.get_form_fields(pdf_5500)
            output_filename = os.path.join(output_dir, f"{row[0].value}_5500.pdf")
        elif gender == "F":
            pdf_template = pdf_5501
            form_field = fillpdfs.get_form_fields(pdf_5501)
            output_filename = os.path.join(output_dir, f"{row[0].value}_5501.pdf")
        else:
            print(f"Warning: Unknown gender '{gender}' at row {row[0].row}. Skipping.")
            continue

        data_dict = {
            "DATE1": current_date,
            "DATE2": current_date,
        }

        data_dict["PREPARED BY"] = prepared_by
        data_dict["PREP_BY_RANK"] = prep_by_rank
        data_dict["APPROVED BY SUPERVISOR"] = approved_by_supervisor
        data_dict["APPR_BY_RANK"] = appr_by_rank

        stop_reading_row = False
        height_weight_pass = False
        exemption = False

        for col_idx, cell in enumerate(row):
            if col_idx == 8 and str(cell.value).strip() == "Pass":
                print(f"Column 9 (index 8) is 'Pass' — stopping further reading for this row at column {col_idx}.")
                stop_reading_row = True
                height_weight_pass = True
                break
            # Check if column 7 ("Yes") and Fail Tape to set special exemption
            elif row[6].value == "Yes" and row[8].value == "Needs Tape":
                print(f"Column 7 (index 6) is 'Yes' and Column 9 (index 8) is Needs Tape — stopping further reading for this row at column {col_idx}.")
                stop_reading_row = True
                exemption = True
                remarks = ACFT_FAIL_HEIGHT_WEIGHT_PASS
                data_dict["REMARKS"] = remarks
                print(f"Writing ACFT_FAIL_HEIGHT_WEIGHT_PASS remark: {remarks}")
            elif row[6].value == "Yes" and row[8].value == "Needs Tape" and row[16].value == "Fail Tape":
                print(f"Column 7 (index 6) is 'Yes' and Column 9 (index 8) is Needs Tape — stopping further reading for this row at column {col_idx}.")
                stop_reading_row = True
                exemption = True
                remarks = ACFT_FAIL_HEIGHT_WEIGHT_PASS
                data_dict["REMARKS"] = remarks
                print(f"Writing ACFT_FAIL_HEIGHT_WEIGHT_PASS remark: {remarks}")

            if cell.value is not None:
                form_field_key = custom_mapping.get(col_idx)
                if form_field_key:
                    if isinstance(form_field_key, list):
                        for key in form_field_key:
                            data_dict[key] = cell.value
                            print(f"Writing '{cell.value}' to form field '{key}'")
                    else:
                        data_dict[form_field_key] = cell.value
                        print(f"Writing '{cell.value}' to form field '{form_field_key}'")
                else:
                    print(f"Warning: No matching form field for column {col_idx} with value '{cell.value}'")

        if exemption and row[8].value == "Needs Tape":
                data_dict["Preparer's Initials"] = preparer_initials
                print(f"Added 'Preparer's Initials': {preparer_initials} because column 6 is 'Yes'")

        if not stop_reading_row:
            body_fat_percentage = float(row[13].value) if row[13].value is not None else 0
            body_fat_standard = float(row[15].value) if row[15].value is not None else 0

            data_dict["BODY FAT PERCENTAGE"] = int(body_fat_percentage * 100)
            data_dict["BODY FAT STANDARD"] = int(body_fat_standard * 100)
            
        else:
            print(f"Skipping body fat and standard values for row {row[0].row} because column 9 is 'Pass'.")



        if height_weight_pass:
            weight = row[5].value  # Column 6 is index 5
            max_weight = row[7].value  # Column 8 is index 7
            if weight is not None and max_weight is not None:
                remarks = HEIGHT_WEIGHT_PASS.format(weight=weight, max_weight=max_weight)
                data_dict["REMARKS"] = remarks
                print(f"Writing HEIGHT_WEIGHT_PASS remark: {remarks}")
            else:
                print(f"Warning: max weight (column 8) missing for {row[0].value}, cannot fill remarks properly.")
        elif row[8].value == "Needs Tape" and row[16].value == "Fail Tape" and exemption == False:
            body_fat_percentage = float(row[13].value) if row[13].value is not None else 0
            body_fat_standard = float(row[15].value) if row[15].value is not None else 0
            data_dict["BODY FAT PERCENTAGE"] = int(body_fat_percentage * 100)
            data_dict["BODY FAT STANDARD"] = int(body_fat_standard * 100)
            remarks = DID_NOT_MEET_STANDARDS.format(
                body_fat_percentage=data_dict["BODY FAT PERCENTAGE"],
                body_fat_standard=data_dict["BODY FAT STANDARD"],
            )
            data_dict["REMARKS"] = remarks
            print(f"Writing DID_NOT_MEET_STANDARDS remark: {remarks}")
        elif row[8].value == "Needs Tape" and row[16].value == "Pass" and exemption == False:
            body_fat_percentage = float(row[13].value) if row[13].value is not None else 0
            body_fat_standard = float(row[15].value) if row[15].value is not None else 0
            data_dict["BODY FAT PERCENTAGE"] = int(body_fat_percentage * 100)
            data_dict["BODY FAT STANDARD"] = int(body_fat_standard * 100)
            remarks = MET_STANDARD.format(
                body_fat_percentage=data_dict["BODY FAT PERCENTAGE"],
                body_fat_standard=data_dict["BODY FAT STANDARD"],
            )
            data_dict["REMARKS"] = remarks
            print(f"Writing MET_STANDARD remark: {remarks}")
        elif row[16].value == "":
            print("The value is blank")
        else:
            print(f"value for cell not PASS and not Fail Tape: '{row[16].value}'")    

        if row[8].value == "Needs Tape" and row[5].value:
            data_dict["WEIGHT_1"] = row[5].value
            print(f"Added 'WEIGHT_1': {row[5].value}")

        if row[8].value == "Pass":
            data_dict["IN COMPLIANCE"] = "Yes"
            print("Added check for compliance: Yes")
        elif row[16].value == "Pass":
            data_dict["IN COMPLIANCE"] = "Yes"
            print("Added check for compliance: Yes")
        elif row[6].value == "Yes" and row[16].value == "Fail Tape":
            data_dict["IN COMPLIANCE"] = "Yes"
            print("Added check for compliance: Yes")
        elif row[16].value == "Fail Tape" and exemption == False:
            data_dict["NOT IN COMPLIANCE"] = "Yes"
            print("Added check for non-compliance: Yes")

        fillpdfs.write_fillable_pdf(
            input_pdf_path=pdf_template,
            output_pdf_path=output_filename,
            data_dict=data_dict,
        )
        print(f"Generated: {output_filename}")

    print("PDF generation complete!")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate PDFs from Excel data and two PDF templates.")
    parser.add_argument("--excel", required=True, help="Path to the Excel file.")
    parser.add_argument("--pdf_5500", required=True, help="Path to the PDF template for Males (Form 5500).")
    parser.add_argument("--pdf_5501", required=True, help="Path to the PDF template for Females (Form 5501).")
    parser.add_argument("--output", required=True, help="Directory to save generated PDFs.")
    parser.add_argument("--date", help="Custom date to use for the DATE1 and DATE2 fields (format: YYYY-MM-DD).")

    args = parser.parse_args()

    custom_date = None
    if args.date:
        try:
            custom_date = datetime.strptime(args.date, "%Y-%m-%d").strftime("%Y-%m-%d")
        except ValueError:
            print("Error: Invalid date format. Use YYYY-MM-DD.")
            exit(1)

    generate_pdfs(
        excel_file=args.excel,
        pdf_5500=args.pdf_5500,
        pdf_5501=args.pdf_5501,
        output_dir=args.output,
        custom_date=custom_date,
    )






# Example:
# python .\generate_pdfs.py --excel "C:\Users\nbing\Desktop\Height_and_Weight/Excel_Master2(TEST).xlsx" --pdf_5500 "C:\Users\nbing\Desktop\Height_and_Weight\BODY_FAT_CONTENT_WORKSHEET_(Male).pdf" --pdf_5501 "C:\Users\nbing\Desktop\Height_and_Weight\BODY_FAT_CONTENT_WORKSHEET_(Female).pdf" --output C:\Users\nbing\Desktop\Height_and_Weight/OUTPUT_TEST
#
# GameDay:
# python .\generate_pdfs.py --excel "C:\Users\nbing\Desktop\Height_and_Weight/Excel_Master2(TEST).xlsx" --pdf_5500 "C:\Users\nbing\Desktop\Height_and_Weight\Forms\BODY_FAT_CONTENT_WORKSHEET_(Male).pdf" --pdf_5501 "C:\Users\nbing\Desktop\Height_and_Weight\Forms\BODY_FAT_CONTENT_WORKSHEET_(Female).pdf" --output C:\Users\nbing\Desktop\Height_and_Weight/OUTPUT_TEST