import streamlit as st
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import altair as alt
import plotly.graph_objects as go
import plotly.io as pio

import xlsxwriter
import base64
import io
import requests
import datetime

from PIL import Image
from streamlit_extras.badges import badge


@st.cache_resource
def get_initial_data(rel_years):
    # Obtaining up-to-date data for application
    api_url = f"https://api.nusmods.com/v2/{rel_years}/moduleInfo.json"
    data = requests.get(api_url).json()
    
    return data


def main():
    col1, col2, col3 = st.columns([0.034, 0.265, 0.035])
    
    with col1:
        url = "https://github.com/tsu2000/nus_gpa_insight/raw/main/images/nus.png"
        response = requests.get(url)
        img = Image.open(io.BytesIO(response.content))
        st.image(img, output_format = "png")

    with col2:
        st.title("&nbsp; NUS GPA Insight")

    with col3:
        badge(type = "github", name = "tsu2000/nus_gpa_insight", url = "https://github.com/tsu2000/nus_gpa_insight")

    # Create sidebar with options
    with st.sidebar:
        st.markdown("# 🛠️ &nbsp; Application Features")
        st.markdown("#####")

        feature = st.radio("Select a feature:", ["Current Course Tracker", 
                                                 "Future GPA Forecast",
                                                 "GPA Calculation Explanation"])   

        st.write("#")
        st.write("##")
        st.write("##")
        st.write("##")

        st.markdown("---")     

        col_a, col_b = st.columns([1.3, 0.9])

        with col_a:
            st.markdown("Data provided by:")
        with col_b:
            url2 = "https://github.com/tsu2000/nus_gpa_insight/raw/main/images/nusmods_banner.png"
            response = requests.get(url2)
            img = Image.open(io.BytesIO(response.content))
            st.image(img, use_container_width = True, output_format = "png")

    # Obtain relevant years for courses
    now = datetime.datetime.now()

    current_year = int(now.strftime("%Y"))
    current_mth_day = now.strftime("%m-%d")

    # Select option
    if feature == "Current Course Tracker":
        calc(current_year, current_mth_day)

    elif feature == "Future GPA Forecast":
        forecast(current_year, current_mth_day)
        
    elif feature == "GPA Calculation Explanation":
        explain()
    
    
def calc(current_year, current_mth_day):
    st.markdown("#### 📝 &nbsp; Current Course Tracker")

    st.markdown("You can add NUS courses to the Course Tracker which can be downloaded to an `.xlsx` file for personal use. You can also view and download statistics about your current GPA based on the added course data from the Course Tracker as a `.pdf` file. _**(Course Info Source: [NUSMods API](https://api.nusmods.com/v2/))**_")

    if current_mth_day < "08-06":
        options = [f"AY {yr-1}/{yr}" for yr in np.arange(2019, current_year+1)]

    elif current_mth_day >= "08-06":
        options = [f"AY {yr}/{yr+1}" for yr in np.arange(2018, current_year+1)]

    if "all_course_data" not in st.session_state:
        st.session_state["all_course_data"] = []

    if "upload_status" not in st.session_state:
        st.session_state["upload_status"] = False

    opt = st.selectbox("Select an Academic Year (AY) to obtain full list of courses available for the time period:", options, index = len(options)-1)

    year_1, year_2 = opt[3:7], opt[8:]
    mod_years = f"{year_1}-{year_2}"

    data = get_initial_data(mod_years)

    grades_to_gpa = {"A+": 5.0,
                     "A": 5.0,
                     "A-": 4.5, 
                     "B+": 4.0, 
                     "B": 3.5, 
                     "B-": 3.0, 
                     "C+": 2.5, 
                     "C": 2.0, 
                     "D+": 1.5, 
                     "D": 1.0, 
                     "F": 0.0, 
                     "S": None, 
                     "U": None,
                     "CS": None,
                     "CU": None,
                     "OVS": None,
                     "OVU": None,
                     "OVI": None,
                     "EXE": None,
                     "IC": None,
                     "IP": None,
                     "W": None}

    cu_dict = {course["moduleCode"]: [course["title"], float(course["moduleCredit"])] for course in data}

    selected_mod = st.selectbox(f"Select a course from AY {year_1}/{year_2} which you have taken from the list (can type to search):", 
                                cu_dict,
                                format_func = lambda key: key + f" - {str(cu_dict[key][0])} [{str(cu_dict[key][1])} CUs]")

    selected_grade = st.selectbox("Select grade you have obtained for the respective course:", grades_to_gpa)

    final_mod_years = mod_years[:4] + "/" + mod_years[5:]

    def results(mod_code, grade):
        mod_title = cu_dict[mod_code][0]
        selected_cus = cu_dict[mod_code][1]
        selected_score = grades_to_gpa[grade]

        return [mod_code, mod_title, selected_cus, grade, selected_score]

    amb_col, rmb_col, clear_col = st.columns([1, 4.2, 0.8]) 

    with amb_col:
        amb = st.button("Add Course")
        if amb:
            st.session_state.all_course_data.append(results(selected_mod, selected_grade) + [final_mod_years])

    with rmb_col:
        rmb = st.button("Remove last row")
        if rmb and st.session_state["all_course_data"] != []:
            st.session_state.all_course_data.remove(st.session_state.all_course_data[-1])

    with clear_col:
        clear = st.button("Clear All")
        if clear:
            st.session_state["all_course_data"] = []

    # Functionality to add mdoules to existing spreadsheet
    upload_xlsx = st.file_uploader("Or, upload an pre-existing `.xlsx` file with course details in the same format:", type = "xlsx", accept_multiple_files = False)

    expected_headers = ["Course Code", "Course Title", "No. of CUs", "Grade", "Grade Points", "AY Taken"]

    if upload_xlsx is not None and st.session_state["upload_status"] == False:
        df_upload = pd.read_excel(upload_xlsx)
        if list(df_upload.columns) != expected_headers:
            st.error("Incorrect column headers. Please use the exact format: " + ", ".join(expected_headers), icon = "🚨")
            st.stop()
        for row in range(len(df_upload)):
            st.session_state.all_course_data.append([i for i in df_upload.iloc[row]])
        st.session_state["upload_status"] = True

    elif upload_xlsx is None:
        st.session_state["upload_status"] = False

    df = pd.DataFrame(columns = expected_headers,
                      data = st.session_state["all_course_data"])
    
    # Change column categories
    df["Grade"] = df["Grade"].astype("category")
    df["Grade"] = pd.Categorical(df["Grade"], categories = list(grades_to_gpa.keys()))

    all_AY = [yr[3:] for yr in options]

    df["AY Taken"] = df["AY Taken"].astype("category")
    df["AY Taken"] = pd.Categorical(df["AY Taken"], categories = all_AY)

    # Show up-to-date dataframe
    st.markdown("###### Add a course and grade to view and download the data table:")

    # Display course data in DataFrame
    if st.session_state["all_course_data"] != []:
        st.dataframe(df.style.format(precision = 1),
                     hide_index = True,
                     use_container_width = True)
        
    analysis_col, export_col = st.columns([1, 0.265]) 

    with export_col:
        def to_excel(df):
            output = io.BytesIO()
            writer = pd.ExcelWriter(output, engine = "xlsxwriter")

            df.to_excel(writer, sheet_name = "course_tracker", index = False)
            workbook = writer.book
            worksheet = writer.sheets["course_tracker"]

            # Add formats and templates here        
            font_color = "#000000"
            header_color = "#ffff00"

            string_template = workbook.add_format(
                {
                    "font_color": font_color, 
                }
            )

            grade_template = workbook.add_format(
                {
                    "font_color": font_color, 
                    "align": "center",
                    "bold": True
                }
            )

            ay_template = workbook.add_format(
                {
                    "font_color": font_color, 
                    "align": "right"
                }
            )

            float_template = workbook.add_format(
                {
                    "num_format": "0.0",
                    "font_color": font_color, 
                }
            )

            header_template = workbook.add_format(
                {
                    "bg_color": header_color, 
                    "border": 1
                }
            )

            column_formats = {
                "A": [string_template, 15],
                "B": [string_template, 50],
                "C": [float_template, 15],
                "D": [grade_template, 15],
                "E": [float_template, 15],
                "F": [ay_template, 15]
            }

            for column in column_formats.keys():
                worksheet.set_column(f"{column}:{column}", column_formats[column][1], column_formats[column][0])
                worksheet.conditional_format(f"{column}1:{column}1", {"type": "no_errors", "format": header_template})

            # Automatically apply Filter function on shape of dataframe
            worksheet.autofilter(0, 0, df.shape[0], df.shape[1]-1)

            # Saving and returning data
            writer.close()
            processed_data = output.getvalue()

            return processed_data

        def get_table_download_link(df):
            """Generates a link allowing the data in a given Pandas DataFrame to be downloaded
            in:  dataframe
            out: href string
            """
            val = to_excel(df)
            b64 = base64.b64encode(val)

            return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="course_details.xlsx">:inbox_tray: Download (.xlsx)</a>' 

        if st.session_state["all_course_data"] != []:
            st.markdown(get_table_download_link(df), unsafe_allow_html = True)

    with analysis_col:
        if st.session_state["all_course_data"] != []:
            analysis = st.button("View Analysis")
        else:
            analysis = None

    if analysis and st.session_state["all_course_data"] != []:

        df2 = df.dropna()
        gpa = sum(df2["No. of CUs"] * df2["Grade Points"]) / sum(df2["No. of CUs"])
        total_cus_gpa = sum(df2["No. of CUs"])

        dp3_gpa = round(gpa, 3)
        dp4_gpa = round(gpa, 4)

        # Degree classification
        if dp3_gpa >= 4.50:
            degree_class = "Honours (Highest Distinction)"
        elif dp3_gpa >= 4.00:
            degree_class = "Honours (Distinction)"
        elif dp3_gpa >= 3.50:
            degree_class = "Honours (Merit)"            
        elif dp3_gpa >= 3.00:
            degree_class = "Honours"        
        elif dp3_gpa >= 2.00:
            degree_class = "Pass"
        else:
            degree_class = "Below Graduation Threshold"

        cus_not_counted = df.loc[(df["Grade"] == "U") | (df["Grade"] == "CU") | (df["Grade"] == "OVU") | (df["Grade"] == "OVI") | (df["Grade"] == "IC") | (df["Grade"] == "IP") | (df["Grade"] == "W")]

        total_completed_cus = sum(df["No. of CUs"]) - sum(cus_not_counted["No. of CUs"])

        # Cumulative courses
        complete_total_mods = len(df)
        conv_mods = len(df2)
        sued_mods = len(df.loc[(df["Grade"] == "U") | (df["Grade"] == "S")])
        cscu_mods = len(df.loc[(df["Grade"] == "CU") | (df["Grade"] == "CS") | (df["Grade"] == "OVU") | (df["Grade"] == "OVS")])
        unrq_mods = len(df.loc[(df["Grade"] == "EXE") | (df["Grade"] == "IC") | (df["Grade"] == "OVI") | (df["Grade"] == "IP") | (df["Grade"] == "W")])

        table_dict = {
            "Final GPA": dp3_gpa,
            "Degree Classification": degree_class,
            "Your GPA (To 4 d.p.)": dp4_gpa,
            "No. of CUs used to calculate GPA": total_cus_gpa,
            "Total No. of CUs completed successfully": total_completed_cus,
            "Total No. of courses attempted (A + B + C + D)": complete_total_mods,
            "No. of courses accounted for in GPA (A)": conv_mods,
            "No. of courses which were S/Ued (B)": sued_mods,
            "No. of CS/CU/OVS/OVU courses taken (C)": cscu_mods,
            "No. of courses with a 'EXE', 'IC', 'OVI', 'IP' or 'W' grade (D)": unrq_mods,
            "Date of Overview": datetime.datetime.now().strftime("%d %b %Y")
        }

        col_fill_colors = ["azure"]*2 + ["lavender"]*3 + ["cornsilk"]*5 + ["honeydew"]
        font_colors = ["mediumblue"]*2 + ["indigo"]*3 + ["saddlebrown"]*5 + ["darkgreen"]

        fig = go.Figure(
            data = [
                go.Table(
                    columnwidth = [2.5, 1.5],
                    header = dict(
                        values = ["<b>Course and GPA Summary Metrics<b>", "<b>Value<b>"],
                        fill_color = "navy",
                        line_color = "black",
                        align = "center",
                        font = dict(color = "white", size = 14, family = "Arial")
                    ),
                    cells = dict(
                        values = [list(table_dict.keys()), list(table_dict.values())], 
                        fill_color = [col_fill_colors, col_fill_colors],
                        line_color = "black",
                        align = ["right", "left"],
                        font = dict(color = [font_colors, font_colors], size = [14, 14], family = "Arial"),
                        height = 25
                    )
                )
            ]
        )

        fig.update_layout(height = 318, width = 700, margin = dict(l = 5, r = 5, t = 5, b = 5))
        st.plotly_chart(fig, use_container_width = True)

        # Create an in-memory buffer
        buffer = io.BytesIO()

        # Save the figure as a pdf to the buffer
        fig.write_image(file = buffer, scale = 6, format = "pdf")

        # Download the pdf from the buffer
        st.download_button(
            label = "Download as PDF",
            data = buffer,
            file_name = "gpa_overview.pdf",
            mime = "application/octet-stream",
            help = "Downloads all course details as a PDF File"
        )

    st.markdown("---")
                       
            
def forecast(current_year, current_mth_day):
    st.markdown("#### 📈 &nbsp; Future GPA Forecast")
    st.markdown("If you provide your current GPA, the number of units used for its calculation (*You can obtain both by using the Current Course Tracker*), and select the courses you plan to take in the upcoming semester which count towards your GPA, you can view the minimum weighted-average unit grades required on all your new courses for you to obtain each classification of honours.")

    cap_col, mc_col = st.columns([1, 1]) 
    
    with cap_col:
        current_gpa = st.number_input('Current GPA to maximum of 4 d.p. (If any):', min_value = 0.0000, max_value = 5.0000, value = 0.0000, step = 0.0001, format = "%0.4f")
        
    with mc_col:
        current_cus = st.number_input('No. of CUs used to calculate current GPA (If any):', min_value = 0.0, max_value = 1000.0, value = 0.0, step = 0.5, format = "%0.1f")

    if current_mth_day < "08-06":
        latest_ay = f"{current_year-1}-{current_year}"
    elif current_mth_day >= "08-06":
        latest_ay = f"{current_year}-{current_year+1}"

    latest_ay_data = get_initial_data(latest_ay)

    cu_latest_dict = {course["moduleCode"]: [course["title"], float(course["moduleCredit"])] for course in latest_ay_data if float(course["moduleCredit"]) != 0}

    future_courses = st.multiselect(f"Select future courses you are planning to take which count towards your GPA (can type to search):", 
                                      cu_latest_dict,
                                      max_selections = 15,
                                      format_func = lambda key: key + f" - {str(cu_latest_dict[key][0])} [{str(cu_latest_dict[key][1])} CUs]")
    
    def future_course_details(mod_code):
        mod_title = cu_latest_dict[mod_code][0]
        selected_cus = cu_latest_dict[mod_code][1]

        return [mod_code, mod_title, selected_cus]
    
    future_data = [future_course_details(code) for code in future_courses]
    
    future_df = pd.DataFrame(future_data, columns = ["Course Code", "Course Title", "No. of CUs"])

    if future_data != []:
    
        st.dataframe(future_df, 
                    hide_index = True,
                    column_config = {
                        "Course Code": st.column_config.Column(width = "small"),
                        "Course Title": st.column_config.Column(width = "large"),
                        "No. of CUs": st.column_config.Column(width = "small")
                        }
                    )
        calc_gpa = st.button("Get GPA Forecast")

        if calc_gpa:
            # Calculate total new CUs from selected courses
            new_cus = future_df["No. of CUs"].sum()
            total_cus = current_cus + new_cus
            
            # Define honours degree classification thresholds
            honours_classes = {
                "🥇 Honours (Highest Distinction)": 4.50,
                "🥈🔼 Honours (Distinction)": 4.00,
                "🥈🔽 Honours (Merit)": 3.50,
                "🥉 Honours": 3.00,
                "🎓 Pass": 2.00
            }

            # Function to calculate the new GPA if all new courses get the same grade
            def calculate_new_gpa_same_grade(grade_point, current_gpa, current_cus, new_cus):
                total_grade_points = current_gpa * current_cus + grade_point * new_cus
                return round(total_grade_points / (current_cus + new_cus), 4) if (current_cus + new_cus) > 0 else 0
            
            # Function to calculate the minimum weighted average GPA for new courses for a target GPA
            def req_weighted_grade_points(target_gpa, current_gpa, current_cus, new_cus):
                new_grade_points_req = target_gpa * (current_cus + new_cus) - current_gpa * current_cus
                new_courses_only_gpa = new_grade_points_req / new_cus
                if new_courses_only_gpa <= 5.0 and new_courses_only_gpa >= 0.0:
                    return round(new_courses_only_gpa, 4)
                else:
                    return "Impossible"
            
            # Function to determine the average letter grades required from the new weighted average GPA
            def points_to_grade_range(req_gpa):
                # Define letter grade to point grade mapping
                grade_gpa_dict = {5.0: "A+/A", 4.5: "A-", 4.0: "B+", 3.5: "B", 3.0: "B-", 2.5: "C+", 2.0: "C", 1.5: "D+", 1.0: "D", 0.0: "F"}
                if req_gpa in grade_gpa_dict.keys():
                    return f"exactly {grade_gpa_dict[req_gpa]}"
                else:
                    sorted_points = sorted(grade_gpa_dict.keys(), reverse=True)

                    for i in range(len(sorted_points) - 1):
                        upper = sorted_points[i]
                        lower = sorted_points[i + 1]
                        if lower < req_gpa < upper:
                            letter1 = grade_gpa_dict[lower]
                            letter2 = grade_gpa_dict[upper]
                            if abs(req_gpa - lower) == abs(req_gpa - upper):
                                return f"exactly between {letter1} and {letter2}"
                            else:
                                closest_letter = letter1 if abs(req_gpa - lower) < abs(req_gpa - upper) else letter2
                                return f"between {letter1} and {letter2}, closer to {closest_letter}"

                    if req_gpa >= max(sorted_points):
                        return f"at least {grade_gpa_dict[max(sorted_points)]}"
                    elif req_gpa <= min(sorted_points):
                        return f"at most {grade_gpa_dict[min(sorted_points)]}"

                    # Exact match fallback
                    return f"exactly {grade_gpa_dict.get(req_gpa, 'Unknown')}"

            new_courses_all_a = calculate_new_gpa_same_grade(5, current_gpa, current_cus, new_cus)
            new_courses_all_f = calculate_new_gpa_same_grade(0, current_gpa, current_cus, new_cus)
    
            # Display course information summary as metrics
            st.markdown("### GPA Forecast Results")

            col1, col2, col3 = st.columns(3)

            with col1:
                st.metric("Current Course Units (CUs)", value = current_cus)
                st.metric("Current GPA", value = current_gpa)
                
            with col2:
                st.metric("New CUs Added", value = new_cus)
                st.metric("Best-Case Scenario (All A+/A)", value = new_courses_all_a, delta =  round(new_courses_all_a - current_gpa, 4))

            with col3:
                st.metric("New Total CUs", value = total_cus)
                st.metric("Worst-Case Scenario (All F)", value = new_courses_all_f, delta = round(new_courses_all_f - current_gpa, 4))

            # Check if user can attain each honours class, if so, list the requirements
            for k, v in honours_classes.items():
                req_new_gpa = req_weighted_grade_points(v, current_gpa, current_cus, new_cus)
                st.markdown(f"#### For {k} - GPA ≥ {v}:")
                if req_new_gpa != "Impossible":
                    st.success("✅ Possible")
                    st.markdown(f"Your unit-weighted average GPA among all new courses must be at least **{req_new_gpa}**")
                    st.markdown(f"Your unit-weighted average grade among all new courses taken should be **{points_to_grade_range(req_new_gpa)}**")
                else:
                    st.error("❌ Impossible to achieve with current GPA and selected courses.")

            st.markdown("---")


def explain():
    st.markdown("#### 🧮 &nbsp; GPA Calculation Explanation")

    st.markdown("To calculate your GPA for $n$ number of relevant courses:")
    
    st.latex(r"""\text{GPA} = \frac{\text{GP}_1 \times{\text{CU}_1} + \text{GP}_2\times{\text{CU}_2} + ... + \text{GP}_n\times{\text{GP}_n}}{\text{CU}_1 + \text{CU}_2 + ... + \text{CU}_n}""")
   
    st.latex(r"""= \sum_{i=1}^{n} \frac{\text{GP}_i \times \text{CU}_i}{\text{CU}_i}""")

    st.markdown(r"""$\text{GP}_i = \text{Course Grade Points for the } i^\text{th} \text{ course used in GPA calculation}$""")
    st.markdown(r"""$\text{CU}_i = \text{Course Units for the } i^\text{th} \text{ course used in GPA calculation}$""")

    st.markdown("---")
    
    st.markdown("Each course taken at NUS usually has a fixed number of Course Units, and each letter grade given after the completion of a course corresponds to a specific number of grade points. NUS uses a 5-point GPA system, with a GPA of 5.0 being the highest possible score. To get your Grade Point Average or GPA, simply do the following:")
                
    st.markdown("1. Obtain the grade points for the course by converting your grade given to the grade points allocated. (E.g. 'A+/A' is 5 grade points, 'B' is 3.5 grade points etc.)*")          
    st.markdown("2. Multiply the grade points you have obtained for each course by the number of course credits assigned to it. (Ignore any courses that do not count toward GPA even if they are requirements for graduation)")
    st.markdown("3. Repeat steps 1 and 2 for all relevant courses.**")
    st.markdown("4. Sum the results of step 3 to get the numerator of the GPA equation.")
    st.markdown("5. Finally, divide the result of step 4 by the total number of course credits (denominator of GPA equation) used to calculate the numerator to get your GPA.")

    st.markdown("Note that NUS rounds to 3 decimal places when determining the honours classification. For example, a GPA of 3.4977 will be rounded to 3.498 and not 3.5.")

    st.markdown("**(*) The mapping between letter grade and grade points for courses that count towards GPA are as follows:**")

    point_grade_mappings = {5.0: "A+/A", 4.5: "A-", 4.0: "B+", 3.5: "B", 3.0: "B-", 2.5: "C+", 2.0: "C", 1.5: "D+", 1.0: "D", 0.0: "F"}
    grade_points, letter_grades = list(point_grade_mappings.keys()), list(point_grade_mappings.values())
    st.table(pd.DataFrame([grade_points], columns = letter_grades, index = ["Grade Points"]).style.format(precision = 1))
    
    st.markdown("**(*\*) The following grades are not factored into GPA, though some of these grades indicate successfully completed course units that count towards degree requirements.**")

    non_gpa_grades_details = {
        "Letter Grade": ["S", "U", "CS", "CU", "OVS", "OVU", "OVI", "EXE", "IC", "IP", "W"],
        "Description": ["Satisfactory", "Unsatisfactory", "Completed Satisfactorily", "Completed Unsatisfactorily", "Overseas Satisfactory", "Overseas Unsatisfactory", "Overseas Incomplete", "Exempted", "Incomplete", "In Progress", "Withdrawn"],
        "Course Units Given": ["Yes", "No", "Yes", "No", "Yes", "No", "No", "Yes", "No", "No", "No"]
    }

    st.table(pd.DataFrame(non_gpa_grades_details))
    
    st.markdown("---")

    st.markdown("Click the link [**here**](https://www.nus.edu.sg/registrar/academic-information-policies/non-graduating/modular-system) to obtain more information about how GPA is calculated at NUS and the relevant grade points for each grade.")
    
    
if __name__ == "__main__":
    st.set_page_config(page_title = "NUS GPA Insight", page_icon = "🧐")
    main()
