
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Path to your template document
template_path = 'D:/Machine Learning/01_notes/CSI Appointment Letter.docx'

data = [
    {'name': 'Aditya Gaikwad', 'designation': 'Technical Head'},
    {'name': 'Kunal Suryawanshi', 'designation': 'Technical Head'},
    {'name': 'Sahil Gite', 'designation': 'Jt-Technical Head'},
    {'name': 'Vedant Shingote', 'designation': 'Jt-Technical Head'},
    {'name': 'Vinay Gadge', 'designation': 'Jt-Technical Head'},

    {'name': 'Rushikesh Ghodke', 'designation': 'Web Master'},
    {'name': 'Pranav Kolhe', 'designation': 'Web Master'},
    {'name': 'Satyajit Borade', 'designation': 'Jt-Web Master'},
    {'name': 'Kanishk Kumar', 'designation': 'Jt-Web Master'},
    {'name': 'Chetashree Bhavsar', 'designation': 'Jt-Web Master'},

    {'name': 'Anushka Gargelwar', 'designation': 'Documentation Head'},
    {'name': 'Suraj Dabhole', 'designation': 'Documentation Head'},
    {'name': 'Pooja Kewat', 'designation': 'Jt-Documentation Head'},
    {'name': 'Bhumi Gambhire', 'designation': 'Jt-Documentation Head'},
    {'name': 'Ananya Bisht', 'designation': 'Jt-Documentation Head'},
    
    {'name': 'Sakshi Khedekar', 'designation': 'Event Head'},
    {'name': 'Harshwardhan Saindane', 'designation': 'Event Head'},
    {'name': 'Samiksha Surwase', 'designation': 'Jt-Event Head'},
    {'name': 'Shadab Mulla', 'designation': 'Jt-Event Head'},
    {'name': 'Bhakti Sarap', 'designation': 'Jt-Event Head'},
    
    {'name': 'Pratik Patil', 'designation': 'Publicity Head'},
    {'name': 'Pranav Pharande', 'designation': 'Publicity Head'},
    {'name': 'Prerna Phadnis', 'designation': 'Jt-Publicity Head'},
    {'name': 'Pranit Govande', 'designation': 'Jt-Publicity Head'},
    {'name': 'Krishna Gandhi', 'designation': 'Jt-Publicity Head'},
    
    {'name': 'Ashish Aher', 'designation': 'Content & Videography Head'},
    {'name': 'Harshita Kadam', 'designation': 'Jt-Content & Videography Head'},
    {'name': 'Sudhanshu Gajbhe', 'designation': 'Jt-Content & Videography Head'},
    {'name': 'Parth Shinde', 'designation': 'Jt-Content & Videography Head'},
    
    {'name': 'Sarthak Kale', 'designation': 'Design Head'},
    {'name': 'Pranav Daware', 'designation': 'Design Head'},
    {'name': 'Jay Rajankar', 'designation': 'Jt- Design Head'},
    {'name': 'Ashita Kolla', 'designation': 'Jt- Design Head'},
    {'name': 'Saburi Kale', 'designation': 'Jt- Design Head'},
]

for entry in data:
    doc = Document(template_path)

    for paragraph in doc.paragraphs:
        if 'NAME_PLACEHOLDER' in paragraph.text:
            for run in paragraph.runs:
                if 'NAME_PLACEHOLDER' in run.text:
                    run.text = run.text.replace('NAME_PLACEHOLDER', entry['name'])
                    run.font.name = 'Calibri '  
                    run.font.size = Pt(24)  
                    run.bold = True  
                    

        if 'DESIGNATION_PLACEHOLDER' in paragraph.text:
            for run in paragraph.runs:
                if 'DESIGNATION_PLACEHOLDER' in run.text:
                    run.text = run.text.replace('DESIGNATION_PLACEHOLDER', entry['designation'])
                    run.font.name = 'Calibri'  
                    run.font.size = Pt(24)  
                    run.bold = True  
                    

    # Save
    output_path = f"D:\\Machine Learning\\01_notes\\{entry['name']}_Appointment_Letter.docx"
    doc.save(output_path)
    print(f"Created document for {entry['name']} at {output_path}")

