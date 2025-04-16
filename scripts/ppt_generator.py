def generate_presentation(excel_path, template_path, output_path, user_name, years_of_service):
    import sys
    import os
    import time
    import pandas as pd
    import comtypes.client
    from comtypes.gen import PowerPoint as PPConst
    
    def rgb_to_ole(red, green, blue):
        """Convert RGB to OLE color format used by PowerPoint."""
        return red + (green * 256) + (blue * 256 * 256)

    def duplicate_slide(presentation, slide_index):
        """Duplicate a slide in a PowerPoint presentation."""
        slides = presentation.Slides
        total_slides = slides.Count
        if slide_index < 1 or slide_index > total_slides:
            raise ValueError(f"Invalid slide index: {slide_index}. Must be between 1 and {total_slides}")
        return slides(slide_index).Duplicate().Item(1)

    def clean_first_slide(presentation):
        """Remove the default 'Employee Name' text from first slide"""
        first_slide = presentation.Slides(1)
        for shape in list(first_slide.Shapes):
            if shape.HasTextFrame and "Employee Name" in shape.TextFrame.TextRange.Text:
                shape.Delete()

    def add_thank_you_slide(presentation):
        """Add 'Thank You' slide preserving both corner designs"""
        thank_you_slide = presentation.Slides(1).Duplicate().Item(1)
        thank_you_slide.MoveTo(presentation.Slides.Count)
        slide_width = presentation.PageSetup.SlideWidth
        slide_height = presentation.PageSetup.SlideHeight
        for shape in list(thank_you_slide.Shapes):
            is_top_left = (shape.Left < 200 and shape.Top < 200) 
            is_bottom_right = (
                shape.Left > slide_width - 400 and 
                shape.Top > slide_height - 200 and 
                shape.Width < 300 and  
                shape.Height < 300
            )
            
            if not (is_top_left or is_bottom_right):
                shape.Delete()
        text_box = thank_you_slide.Shapes.AddTextbox(
            1, 
            slide_width/2 - 150,  
            slide_height/2 - 40,  
            300, 80
        )
        text_frame = text_box.TextFrame.TextRange
        text_frame.Text = "Thank you"
        text_frame.Font.Size = 65
        text_frame.Font.Name = "Century Gothic"
        text_frame.Font.Color.RGB = rgb_to_ole(255, 0, 0)
        text_frame.Font.Bold = True
        text_frame.Font.Italic = True
        text_frame.ParagraphFormat.Alignment = PPConst.ppAlignCenter
        text_box.TextFrame.WordWrap = False

    def create_message_template_slide(presentation):
        """Create message template by duplicating first slide and removing anniversary logo + name"""
        message_slide = presentation.Slides(1).Duplicate().Item(1)
        for shape in list(message_slide.Shapes):
            if shape.HasTextFrame and "Employee Name" in shape.TextFrame.TextRange.Text:
                shape.Delete()
            elif shape.Type == 13 and not (shape.Left < 100 and shape.Top < 100):
                shape.Delete()
        
        return message_slide

    def modify_template(input_path, output_path, years_of_service):
        """Process template - keeps anniversary logo+name on first slide only"""
        try:
            powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
            powerpoint.Visible = 1
            presentation = powerpoint.Presentations.Open(os.path.abspath(input_path))
            
            
            selected_anniversary = min(max(1, int(years_of_service)), 4)
            print(f"Selected anniversary template: {selected_anniversary} year(s)")
            
            
            for i in reversed(range(1, presentation.Slides.Count + 1)):
                if i != selected_anniversary:
                    presentation.Slides(i).Delete()
                    
            
            create_message_template_slide(presentation)
            
            
            temp_path = output_path.replace(".pptx", "_temp.pptx")
            presentation.SaveAs(os.path.abspath(temp_path))
            presentation.Close()
            time.sleep(1)
            
            
            if os.path.exists(output_path):
                try:
                    os.remove(output_path)
                except PermissionError:
                    print(f"Please close: {output_path}")
                    return False
                    
            os.replace(temp_path, output_path)
            print(f"Template ready: {output_path}")
            return True
            
        except Exception as e:
            print(f"Error modifying template: {e}")
            return False
        finally:
            try:
                powerpoint.Quit()
            except:
                pass

    def replace_text_in_slide(slide, replacements):
        """Replace placeholders in a slide's text shapes."""
        for shape in slide.Shapes:
            if shape.HasTextFrame:
                text_range = shape.TextFrame.TextRange
                text = text_range.Text
                for key, value in replacements.items():
                    if key.lower() in text.lower():
                        text = text.replace(key, value)
                        text = text.replace(key.upper(), value)
                        text = text.replace(key.title(), value)
                text_range.Text = text

    
    print("Step 1: Modifying template...")
    modified_template_path = os.path.splitext(template_path)[0] + "_modified.pptx"
    if not modify_template(template_path, modified_template_path, years_of_service):
        print("Failed to modify template. Exiting.")
        return
    
    
    time.sleep(2)
    
    
    print("Step 2: Generating presentation...")
    if not os.path.exists(modified_template_path):
        print(f"Error: Modified template '{modified_template_path}' not found.")
        return

    try:
        
        df = pd.read_excel(excel_path).rename(columns=lambda x: x.strip())
        required_columns = ['Name', 'Wishes']
        if not all(col in df.columns for col in required_columns):
            print(f"Error: Missing required columns in Excel file. Found: {df.columns}")
            return
        
        
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = 1
        presentation = powerpoint.Presentations.Open(os.path.abspath(modified_template_path))
        
        
        clean_first_slide(presentation)
        first_slide = presentation.Slides(1)
        

        name_left_position = first_slide.Master.Width - 450  
        name_top_position = first_slide.Master.Height - 100  
        name_text_box = first_slide.Shapes.AddTextbox(1, name_left_position, name_top_position, 250, 50)
        name_text_frame = name_text_box.TextFrame.TextRange
        name_text_frame.Text = user_name
        name_text_frame.Font.Size = 24
        name_text_frame.Font.Name = "Century Gothic"
        name_text_frame.Font.Color.RGB = rgb_to_ole(0, 0, 0) 
        name_text_frame.Font.Bold = True
        name_text_frame.ParagraphFormat.Alignment = PPConst.ppAlignCenter
        
        
        years_left_position = first_slide.Master.Width - 450
        years_top_position = first_slide.Master.Height - 150
        years_text_box = first_slide.Shapes.AddTextbox(1, years_left_position, years_top_position, 250, 50)
        years_text_frame = years_text_box.TextFrame.TextRange
        years_text_frame.Text = f"{years_of_service} Years of Service"
        years_text_frame.Font.Size = 20
        years_text_frame.Font.Name = "Century Gothic"
        years_text_frame.Font.Color.RGB = rgb_to_ole(255, 0, 0)
        years_text_frame.Font.Bold = True
        years_text_frame.ParagraphFormat.Alignment = PPConst.ppAlignCenter

        if presentation.Slides.Count < 2:
            print("Template must have at least 2 slides (cover + message template).")
            presentation.Close()
            powerpoint.Quit()
            return
    
        message_template_index = 2  
        max_messages_per_slide = 2  
        messages_on_slide = 0  
        current_slide = None  
        text_positions_two = [(100, 100), (100, 250)]
        text_positions_one = [(presentation.PageSetup.SlideWidth // 2 - 250, 
                      presentation.PageSetup.SlideHeight // 3)]
        df = df.iloc[::-1] 

        for index, row in df.iterrows():
            name = " ".join(word.capitalize() for word in str(row['Name']).strip().split() if word)
            message = str(row['Wishes']).strip()

            if not message or message.lower() == 'nan':
                continue  
            is_long_message = len(message) > 150
            if is_long_message or messages_on_slide >= max_messages_per_slide or current_slide is None:
                current_slide = duplicate_slide(presentation, message_template_index)
                messages_on_slide = 0  
            text_positions = text_positions_one if is_long_message else text_positions_two
            text_position = text_positions[messages_on_slide]  
            text_box = current_slide.Shapes.AddTextbox(1, *text_position, 500, 100)
            text_frame = text_box.TextFrame.TextRange
            text_frame.Text = message
            text_frame.Font.Size = 18
            text_frame.Font.Name = "Century Gothic"
            text_frame.Font.Bold = False
            text_frame.Font.Color.RGB = rgb_to_ole(0, 0, 0)
            text_frame.ParagraphFormat.Alignment = PPConst.ppAlignJustify
            message_height = text_box.TextFrame.TextRange.BoundHeight  
            signature_left = text_position[0]  
            signature_top = text_position[1] + message_height + 10  
            signature_box = current_slide.Shapes.AddTextbox(1, signature_left, signature_top, 500, 30)  
            signature_frame = signature_box.TextFrame.TextRange
            signature_frame.Text = f"- {', '.join(name.splitlines())}"  
            signature_frame.Font.Size = 18
            signature_frame.Font.Name = "Century Gothic"
            signature_frame.Font.Color.RGB = rgb_to_ole(255, 0, 0) 
            signature_frame.Font.Bold = True
            signature_frame.ParagraphFormat.Alignment = PPConst.ppAlignRight  
            signature_box.TextFrame.WordWrap = False  
            messages_on_slide += 1 if not is_long_message else max_messages_per_slide  
        
        
        if presentation.Slides.Count > 2:
            presentation.Slides(2).Delete()  

        
        add_thank_you_slide(presentation)

        
        presentation.SaveAs(os.path.abspath(output_path))
        presentation.Close()
        powerpoint.Quit()
        
        
        try:
            os.remove(modified_template_path)
        except:
            pass
            
        print(f"Presentation created successfully: {output_path}")

    except Exception as e:
        print(f"Error in generate_presentation: {e}")
        try:
            presentation.Close()
            powerpoint.Quit()
        except:
            pass