import csv
import fitz
import tkinter as tk
from tkinter import filedialog
import os
import math
from shapely.geometry import LineString, Point
from collections import Counter

def get_csv_values(csv_path):
    values = []
    with open(csv_path, mode='r', newline='', encoding='utf-8') as file:
        reader = csv.reader(file)
        next(reader)  # skip header
        for row in reader:
            values.append(row[0])
    return values

def most_common_value(values):
    # Count the occurrences of each value in the list
    counter = Counter(values)
    
    # Find the most common value(s)
    most_common_values = counter.most_common(1)
    
    # Return the most common value(s)
    return most_common_values[0][0] if most_common_values else None


def find_matches_in_pdf(pdf_path, search_terms):
    matched_terms_and_quads = []
    doc = fitz.open(pdf_path)
    pages = [6]
    for page_num in range(doc.page_count):
        page = doc[page_num]
        #print(f"Searching page{page_num}")
        for term in tqdm(search_terms, desc=f"Finding Suspected Beams on  page: {page_num}", unit="Size"):
            quads = page.search_for(term, quads=True)
            
            if quads:
                for quad in quads:
                    matched_terms_and_quads.append((term, page_num, quad))
                    
    doc.close()
    #print(matched_terms_and_quads)
    return matched_terms_and_quads

import math
from tqdm import tqdm

def find_scale_in_pdf(pdf_path):
    scales = []
    all_scales = ['Scale:','SCALE']
    doc = fitz.open(pdf_path)
    cat = doc.pdf_catalog() 
    print(doc.xref_object(cat))
    # bsi_definition = (doc.xref_get_key(cat, "BSIAnnotColumns"))
    # bsi_xref=int(bsi_definition[1][0])
    # bsi_columns=(doc.xref_object(bsi_xref))
    # parsed_columns=parse_pdf_dict(bsi_columns)
#     for d in parsed_columns:
#         print(d['/Name'])
#         for key in d:
#                 print(f"Key: {key}, Value: {d[key]}")
#                 #The value stored in /Items represents the xref of the list of possible values offset by +1

#     for page in doc:
#         blocks = page.get_text("blocks")
#         for block in blocks:
#             _, _, _, _, text, = block[:5]
#             if any(val in text for val in all_scales):
#                 scales.append((page.number, block))
#     return scales 
# import re

def parse_pdf_dict(pdf_dict_str):
    """
    Parses a string representing PDF dictionary objects into a list of Python dictionaries.

    :param pdf_dict_str: A string containing PDF dictionary objects.
    :return: A list of Python dictionaries representing the parsed data.
    """

    # Use regular expression to find all instances of dictionary structures within the string.
    # Each dictionary is enclosed in '<<' and '>>'.
    dicts = re.findall(r'<<.*?>>', pdf_dict_str, re.DOTALL)

    parsed_data = []

    for d in dicts:
        # Remove the '<<' and '>>' from the beginning and end of each dictionary string.
        d = d[2:-2].strip()

        # Split the dictionary string into key-value pairs.
        # This regex matches patterns like '/Key [array]' or '/Key (value)' or '/Key value'.
        pairs = re.findall(r'/\w+ \[.*?\]|/\w+ \(.*?\)|/\w+ \d+', d)

        # Initialize an empty dictionary to store the parsed key-value pairs.
        dict_obj = {}
        for pair in pairs:
            # Split each pair into key and value at the first space.
            key, value = pair.split(' ', 1)
            
            # Remove parentheses or brackets from value, if present.
            value = value.strip("()[]")

            # Convert numerical values from string to integers.
            # This assumes that all numerical values are integers.
            if value.isdigit():
                value = int(value)

            # Assign the key-value pair to our dictionary.
            dict_obj[key] = value

        # Append the parsed dictionary to our list.
        parsed_data.append(dict_obj)

    return parsed_data


def get_angle(p1, p2):
    x1, y1 = p1
    x2, y2 = p2
    return math.degrees(math.atan2(y2 - y1, x2 - x1))

from tqdm import tqdm

def extend_line(p1, p2, extension):
    # Calculate the vector from p1 to p2
    vector = [p2[0] - p1[0], p2[1] - p1[1]]

    # Calculate the length of the vector
    length = math.sqrt(vector[0]**2 + vector[1]**2)

    # Calculate the unit vector
    unit_vector = [vector[0] / length, vector[1] / length]

    # Extend the line by adding and subtracting the unit vector scaled by the extension amount
    extended_p1 = [p1[0] - unit_vector[0] * extension, p1[1] - unit_vector[1] * extension]
    extended_p2 = [p2[0] + unit_vector[0] * extension, p2[1] + unit_vector[1] * extension]

    return extended_p1, extended_p2



# Function to filter lines by length
def filter_lines(drawings,length_threshold,width_threshold=0):
    filtered_lines = []
    
    length_threshold= 24
    drawing_list=drawings.get_drawings()
    #print(width_threshold)
    for item in drawing_list:
        #print(item)


        if item['width']:
            try:
                if item['items'][0][0] == 'l' and item['width']>=width_threshold:


                    for subitem in item['items']:

                        if subitem[0]=='l' :
                            line_type, p0, p1 = subitem
                            lx0, ly0 = p0
                            lx1, ly1 = p1
                            line_length = math.hypot(lx1 - lx0, ly1 - ly0)
                            line_width = item['width']
                            if line_length> length_threshold:
                                filtered_lines.append((line_type, p0, p1, line_width))
                                #print(f"Line Width={line_width}")
                        else:

                            pass
                    
            except ValueError as e:
                print(f"Unexpected data format: {item['items'][0]}, error: {str(e)}")
    return filtered_lines

# Function to find the closest line to a given text block accepts a width and length threshold 
def find_closest_line(drawings, text_block,found_text, width_threshold=0, length_threshold=0):
    # Create a filtered list of lines
    filtered_lines = filter_lines(drawings,length_threshold,width_threshold)
    

    print(filtered_lines)
    x0, y0, x1, y1 = text_block[:4]
    text_ll = x0,y0
    text_lr = x1,y1

    text_center = ((x0 + x1) / 2, (y0 + y1) / 2)
    # Initialize variables for tracking the closest line
    closest_line = None
    min_distance = float('inf')

    

    for ltype,p0,p1,width in tqdm(filtered_lines, desc=f"Analyzing Vector Geometry for {found_text}", unit="line"):
        try:
            line_type=ltype
            lx0, ly0 = p0
            lx1, ly1 = p1
            line_center = ((lx0 + lx1) / 2, (ly0 + ly1) / 2)
            
            # Calculate Euclidean distance
            distance = ((text_center[0] - line_center[0]) ** 2 +
                        (text_center[1] - line_center[1]) ** 2) ** 0.5
            
            #Check that the rotation of the text is the same as the rotation of the line

            if abs(get_angle(p0,p1) - get_angle(text_ll,text_lr))<2 or abs(get_angle(p1,p0) - get_angle(text_ll,text_lr))<2  :

                if distance < min_distance:
                    min_distance = distance
                    closest_line = (lx0, ly0, lx1, ly1,text_center, distance,found_text,width)
            



        except ValueError as e:
            print(f"Unexpected data format: {item['items'][0]}, error: {str(e)}")


    return closest_line


def closest_distance(point1, point2):
    return math.sqrt((point2[0] - point1[0]) ** 2 + (point2[1] - point1[1]) ** 2)

def closest_point(primary_point, other_points):
    min_distance = float('inf')
    closest_point = None
    
    for point in other_points:
        dist = closest_distance(primary_point, point)
        if dist < min_distance:
            min_distance = dist
            closest_point = point
    
    return closest_point


def distance(point1, point2):
    return ((point1[0] - point2[0]) ** 2 + (point1[1] - point2[1]) ** 2) ** 0.5

def is_min_distance(point1,point2):
    if point1 - point2 < 1.5:
        return True
    else:
        return False


def is_point_on_line_segment(line_point1, line_point2, test_point):
    # Calculate length of the line segment
    line_length = distance(line_point1, line_point2)
    
    # Calculate distances from the line points to the test point
    distance_from_point1 = distance(line_point1, test_point)
    distance_from_point2 = distance(line_point2, test_point)
    
    # Check if the sum of distances is approximately equal to the length of the line segment
    return math.isclose(distance_from_point1 + distance_from_point2, line_length, rel_tol=1e-9)

def beam_line_intersection(current_line,beam_lines):  
        
    extended_line = extend_line(current_line[0],current_line[1],50)
    
    #Find intersection for P0 endpoint
    
    p0_all_intersection = []  # all intersection point
    p0_seg_intersection = []  # intersection within the extended line segment
    p0_intersection = None    # Final intersection point we needed
    
    for i in beam_lines[0]:
        line1 = LineString(extended_line)
        line2 = LineString(i)
        extended_line2= extend_line(i[0],i[1],50)
        line2 = LineString(extended_line2)
        intersection = line1.intersection(line2)

        if intersection:      
            try:
                intersection_coords = intersection.coords
                intersection_x0 = intersection_coords[0][0]
                intersection_y0 = intersection_coords[0][1]
                if (intersection_x0,intersection_y0) != current_line[0]:
                    p0_all_intersection.append((intersection_x0,intersection_y0))
                
            except Exception as e:
                print('Error - in beam_line_intersection : ',e )
    
    if len(p0_all_intersection) > 0 :
        for i in p0_all_intersection:      
            if is_point_on_line_segment(current_line[0],extended_line[0],i):
                p0_seg_intersection.append(i)
        if len(p0_seg_intersection) > 0:
            p0_intersection = closest_point(current_line[0],p0_seg_intersection)
        else:
            p0_intersection = current_line[0]                        
    else:
        p0_intersection = current_line[0]
        
    #Find intersection for P1 endpoint
        
    p1_all_intersection = []
    p1_seg_intersection = []
    p1_intersection = None
    for i in beam_lines[0]:
        line1 = LineString(extended_line)
        line2 = LineString(i)
        intersection = line1.intersection(line2)
        
        if intersection:
            try:
                intersection_coords = intersection.coords
                intersection_x0 = intersection_coords[0][0]
                intersection_y0 = intersection_coords[0][1]
                if (intersection_x0,intersection_y0) != current_line[1]:
                    p1_all_intersection.append((intersection_x0,intersection_y0))
            except Exception as e:
                print('Error - in beam_line_intersection : ',e )
            
    
    if len(p1_all_intersection) > 0 :
        for i in p1_all_intersection:
            if is_point_on_line_segment(current_line[1],extended_line[1],i):
                p1_seg_intersection.append(i)
        if len(p1_seg_intersection) > 0:
            p1_intersection = closest_point(current_line[1],p1_seg_intersection)
        else:
            p1_intersection = current_line[1]
                            
    else:
        p1_intersection = current_line[1]
        
    if p0_intersection == None:
        p0_intersection = current_line[0]
    if p1_intersection == None:
        p1_intersection = current_line[1]

    return (p0_intersection,p1_intersection)

def find_beam_lines(pdf_path, matches,width_threshold=0):
    beam_lines = {}
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"Error opening PDF: {e}")
        return

    widths= []

    for text, page_num, quad in matches:

        page = doc[page_num]
        try:

            p1,p2,p3,p4 = quad
            x1, y1 = p1
            x2, y2 = p2
            x3, y3 = p3
            x4, y4 = p4
            
            text_bottom = p3[0],p3[1],p4[0],p4[1]        
            closest_line = find_closest_line(page, text_bottom, text, width_threshold)
            
            if closest_line == None:
                print('Closest_line is Null')   
            if closest_line:
                
                cl_x0, cl_y0, cl_x1, cl_y1,text_center, distance,found_text, line_width = closest_line   
                widths.append(line_width)
                #print(line_width)     
                current_line = ((cl_x0, cl_y0), (cl_x1, cl_y1))
                if page_num in beam_lines:
                    beam_lines[page_num].append(current_line)
                else:
                    beam_lines[page_num] = [current_line]

        except Exception as e:
            print(f"Error processing match on page {page_num}: {e}")          
    if width_threshold == 0:
        most_common_width= most_common_value(widths)
        find_beam_lines(pdf_path,matches,most_common_width)

    print('beam_lines are ', beam_lines)
    return beam_lines

def annotate_matches_in_pdf(pdf_path, matches, output_path):
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        print(f"Error opening PDF: {e}")
        return
    
    beam_lines= find_beam_lines(pdf_path, matches)

    for text, page_num, quad in matches:
        page = doc[page_num]

        try:
            # Extracting the points from the Quad object
            #print(quad)
            p1,p2,p3,p4 = quad

            # Extracting x, y coordinates from Point objects
            x1, y1 = p1
            x2, y2 = p2
            x3, y3 = p3
            x4, y4 = p4
            
            # Adding a highlight annotation using the quad points
            #highlight=page.add_highlight_annot([x1, y1, x2, y2, x3, y3, x4, y4])
            highlight = page.add_highlight_annot(quad)
            


            #highlight.set_info(fitz.PDF_INFO_ANNOTATION, bsi_data)
            
                        

            text_bottom = p3[0],p3[1],p4[0],p4[1]
            
            closest_line = find_closest_line(page, text_bottom, text)
            
            
            if closest_line:
                #print(closest_line[7])
                cl_x0, cl_y0, cl_x1, cl_y1, text_center, distance, found_text,width = closest_line
                current_line = ((cl_x0, cl_y0), (cl_x1, cl_y1))
                p1, p2 = beam_line_intersection(current_line, beam_lines)
                line_annot = page.add_line_annot(p1, p2)
                line_annot.set_info(title="Beam", subject=text)
                if distance < 50:
                    line_annot.set_colors({"stroke": (0, 1, 0), "title": (0, 0, 0)})  # Green line for close matches
                    highlight.set_colors(stroke=(1, 1, 0))  # Yellow highlight
                else:
                    line_annot.set_colors({"stroke": (1, 0, 0), "title": (0, 0, 0)})  # Red line for far matches
                    highlight.set_colors(stroke=(1, 0, 0))  # Red highlight
                    leader_annot=page.add_line_annot(quad.ll,p1)
                    leader_annot.update()
                line_annot.update()
                highlight.update()
                doc.xref_set_key(line_annot.xref,"BSIColumnData", f"[(Beam)()()({text})()(1)()()()()()()()()()(No)()(0)(No)()]")
        except Exception as e:
            print(f"Error processing match on page {page_num}: {e} ",e.args)

            

    try:
        doc.save(output_path, garbage=4)
        doc.close()
        print("Annotation complete. Output saved to:", output_path)
    except Exception as e:
        print(f"Error saving annotated PDF: {e}")


def save_last_folder(folder_path):
    with open("last_folder.txt", "w") as file:
        file.write(folder_path)

def get_last_folder():
    try:
        with open("last_folder.txt", "r") as file:
            return file.read()
    except FileNotFoundError:
        return ""

def main():
    root = tk.Tk()
    root.withdraw()  # We don't want a full GUI, so keep the root window from appearing
    
    last_folder = get_last_folder()
    
    csv_msg = "Select the CSV file"
    #csv_path = filedialog.askopenfilename(initialdir=last_folder, title=csv_msg, filetypes=[("CSV files", "*.csv")])
    csv_path = "I:\\Shared drives\\United Structural\\Code Library\\FindBeamSizes\\BluebeamShapes.csv"
    #output_csv_path=filedialog.asksaveasfilename(initialdir=last_folder,defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    found_beams = []
    if not csv_path:
        print("No CSV file selected. Exiting...")
        return
    
    save_last_folder(os.path.dirname(csv_path))

    pdf_msg = "Select the PDF file"
    pdf_path = filedialog.askopenfilename(initialdir=last_folder, title=pdf_msg, filetypes=[("PDF files", "*.pdf")])
    
    if not pdf_path:
        print("No PDF file selected. Exiting...")
        return
    
    save_last_folder(os.path.dirname(pdf_path))

    output_pdf_path = filedialog.asksaveasfilename(defaultextension=".pdf", initialdir=last_folder, title="Save the output PDF as", filetypes=[("PDF files", "*.pdf")])
    
    if not output_pdf_path:
        print("No output file specified. Exiting...")
        return
    
    save_last_folder(os.path.dirname(output_pdf_path))
    
    csv_values = get_csv_values(csv_path)
    matches = find_matches_in_pdf(pdf_path, csv_values)

    #scales = find_scale_in_pdf(pdf_path)
    #print(scales)
    annotate_matches_in_pdf(pdf_path, matches, output_pdf_path)

if __name__ == "__main__":
    main()
