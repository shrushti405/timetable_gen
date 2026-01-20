import pandas as pd
from datetime import datetime, timedelta
import json
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import random
import traceback
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

class TimetableGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Dynamic Timetable Generator - Multi-Teacher Support")
        self.root.geometry("1400x900")
        
        # Initialize data structures
        self.timetable_data = {}
        self.school_timing = {}
        self.subjects = []
        self.standards = []
        self.sections = []
        self.teachers = []
        self.subject_teachers = {}  # subject -> [teacher1, teacher2, ...] (CHANGED)
        self.standard_subjects = {}  # standard -> list of subjects
        self.teacher_hours = {}  # teacher -> {'min_hours': X, 'max_hours': Y}
        self.config = {}
        self.days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create tabs
        self.create_welcome_tab()
        self.create_school_timing_tab()
        self.create_periods_tab()
        self.create_subjects_tab()
        self.create_classes_tab()
        self.create_teachers_tab()
        self.create_assignments_tab()
        self.create_generate_tab()
        self.create_view_tab()
        self.create_export_tab()
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
    def create_welcome_tab(self):
        """Create welcome/introduction tab"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Welcome")
        
        # Title
        title = ttk.Label(tab, text="🎓 Advanced Timetable Generator", 
                         font=('Arial', 24, 'bold'))
        title.pack(pady=30)
        
        # Description
        desc = ttk.Label(tab, text="Create conflict-free school timetables with multiple teachers per subject\n\n"
                                  "Follow the steps in order from left to right:", 
                        font=('Arial', 12))
        desc.pack(pady=20)
        
        # Steps
        steps = [
            "1. School Timing - Set school hours and working days",
            "2. Periods - Configure period duration and recess times",
            "3. Subjects - Add all subjects to be taught",
            "4. Classes - Add standards/classes and sections",
            "5. Teachers - Add teaching staff",
            "6. Assignments - Assign multiple teachers to subjects",
            "7. Generate - Create optimized conflict-free timetable",
            "8. View - Preview generated timetables",
            "9. Export - Save to Excel or JSON"
        ]
        
        for step in steps:
            lbl = ttk.Label(tab, text=step, font=('Arial', 11))
            lbl.pack(pady=5)
        
        # Quick actions frame
        quick_frame = ttk.LabelFrame(tab, text="Quick Actions", padding=20)
        quick_frame.pack(pady=30, padx=50, fill='x')
        
        ttk.Button(quick_frame, text="📂 Load Configuration", 
                  command=self.load_configuration).pack(side=tk.LEFT, padx=10)
        ttk.Button(quick_frame, text="💾 Save Configuration", 
                  command=self.save_configuration).pack(side=tk.LEFT, padx=10)
        ttk.Button(quick_frame, text="❓ Help", 
                  command=self.show_help).pack(side=tk.RIGHT, padx=10)
    
    def create_school_timing_tab(self):
        """Create school timing configuration tab"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="School Timing")
        
        # Title
        title = ttk.Label(tab, text="School Timing Configuration", 
                         font=('Arial', 16, 'bold'))
        title.pack(pady=20)
        
        # Input frame
        input_frame = ttk.Frame(tab)
        input_frame.pack(pady=20, padx=50)
        
        # Start time
        ttk.Label(input_frame, text="School Start Time (HH:MM):", 
                 font=('Arial', 11)).grid(row=0, column=0, sticky=tk.W, pady=10)
        self.start_time_var = tk.StringVar(value="07:30")
        ttk.Entry(input_frame, textvariable=self.start_time_var, 
                 font=('Arial', 11), width=15).grid(row=0, column=1, pady=10, padx=10)
        
        # End time
        ttk.Label(input_frame, text="School End Time (HH:MM):", 
                 font=('Arial', 11)).grid(row=1, column=0, sticky=tk.W, pady=10)
        self.end_time_var = tk.StringVar(value="12:30")
        ttk.Entry(input_frame, textvariable=self.end_time_var, 
                 font=('Arial', 11), width=15).grid(row=1, column=1, pady=10, padx=10)
        
        # Days of operation
        ttk.Label(input_frame, text="Working Days:", 
                 font=('Arial', 11)).grid(row=2, column=0, sticky=tk.W, pady=10)
        
        days_frame = ttk.Frame(input_frame)
        days_frame.grid(row=2, column=1, sticky=tk.W, pady=10)
        
        self.day_vars = {}
        days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
        for i, day in enumerate(days):
            var = tk.BooleanVar(value=True if i < 5 else False)
            self.day_vars[day] = var
            chk = ttk.Checkbutton(days_frame, text=day, variable=var)
            chk.grid(row=i//3, column=i%3, sticky=tk.W, padx=5)
        
        # Save button
        ttk.Button(tab, text="Save School Timing", 
                  command=self.save_school_timing, 
                  style='Accent.TButton').pack(pady=30)
        
        # Preview label
        self.timing_preview = ttk.Label(tab, text="", font=('Arial', 10))
        self.timing_preview.pack(pady=10)
    
    def create_periods_tab(self):
        """Create periods configuration tab"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Periods")
        
        # Title
        title = ttk.Label(tab, text="Periods and Recess Configuration", 
                         font=('Arial', 16, 'bold'))
        title.pack(pady=20)
        
        # Period configuration frame
        period_frame = ttk.LabelFrame(tab, text="Period Settings", padding=20)
        period_frame.pack(pady=10, padx=50, fill='x')
        
        # Period duration
        ttk.Label(period_frame, text="Period Duration (minutes):", 
                 font=('Arial', 11)).grid(row=0, column=0, sticky=tk.W, pady=10)
        self.period_duration_var = tk.IntVar(value=35)
        ttk.Entry(period_frame, textvariable=self.period_duration_var, 
                 font=('Arial', 11), width=15).grid(row=0, column=1, pady=10, padx=10)
        
        # Period overtime
        ttk.Label(period_frame, text="Period Overtime (minutes):", 
                 font=('Arial', 11)).grid(row=1, column=0, sticky=tk.W, pady=10)
        self.period_overtime_var = tk.IntVar(value=0)
        ttk.Entry(period_frame, textvariable=self.period_overtime_var, 
                 font=('Arial', 11), width=15).grid(row=1, column=1, pady=10, padx=10)
        
        # Teaching hours
        ttk.Label(period_frame, text="Teaching Hours per Day:", 
                 font=('Arial', 11)).grid(row=2, column=0, sticky=tk.W, pady=10)
        self.teaching_hours_var = tk.IntVar(value=5)
        ttk.Entry(period_frame, textvariable=self.teaching_hours_var, 
                 font=('Arial', 11), width=15).grid(row=2, column=1, pady=10, padx=10)
        
        # Recess configuration frame
        recess_frame = ttk.LabelFrame(tab, text="Recess Settings", padding=20)
        recess_frame.pack(pady=20, padx=50, fill='x')
        
        # Recess list
        recess_list_frame = ttk.Frame(recess_frame)
        recess_list_frame.pack(fill='both', expand=True)
        
        # Create treeview for recess
        columns = ('Name', 'Start Time', 'Duration')
        self.recess_tree = ttk.Treeview(recess_list_frame, columns=columns, 
                                       show='headings', height=4)
        
        for col in columns:
            self.recess_tree.heading(col, text=col)
            self.recess_tree.column(col, width=150)
        
        self.recess_tree.pack(side=tk.LEFT, fill='both', expand=True)
        
        # Scrollbar for recess tree
        scrollbar = ttk.Scrollbar(recess_list_frame, orient=tk.VERTICAL, 
                                 command=self.recess_tree.yview)
        self.recess_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Add/Edit recess frame
        add_recess_frame = ttk.Frame(recess_frame)
        add_recess_frame.pack(pady=10, fill='x')
        
        ttk.Label(add_recess_frame, text="Name:").pack(side=tk.LEFT, padx=5)
        self.recess_name_var = tk.StringVar(value="Recess")
        ttk.Entry(add_recess_frame, textvariable=self.recess_name_var, 
                 width=15).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(add_recess_frame, text="Start Time:").pack(side=tk.LEFT, padx=5)
        self.recess_time_var = tk.StringVar(value="11:30")
        ttk.Entry(add_recess_frame, textvariable=self.recess_time_var, 
                 width=10).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(add_recess_frame, text="Duration:").pack(side=tk.LEFT, padx=5)
        self.recess_duration_var = tk.IntVar(value=30)
        ttk.Entry(add_recess_frame, textvariable=self.recess_duration_var, 
                 width=10).pack(side=tk.LEFT, padx=5)
        
        # Recess buttons
        button_frame = ttk.Frame(recess_frame)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="Add Recess", 
                  command=self.add_recess).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Remove Selected", 
                  command=self.remove_recess).pack(side=tk.LEFT, padx=5)
        
        # Save button
        ttk.Button(tab, text="Save Period Settings", 
                  command=self.save_period_settings, 
                  style='Accent.TButton').pack(pady=30)
    
    def create_subjects_tab(self):
        """Create subjects management tab"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Subjects")
        
        # Title
        title = ttk.Label(tab, text="Subjects Management", 
                         font=('Arial', 16, 'bold'))
        title.pack(pady=20)
        
        # Add subject frame
        add_frame = ttk.Frame(tab)
        add_frame.pack(pady=10, padx=50, fill='x')
        
        ttk.Label(add_frame, text="Subject Name:", 
                 font=('Arial', 11)).pack(side=tk.LEFT)
        self.subject_entry_var = tk.StringVar()
        subject_entry = ttk.Entry(add_frame, textvariable=self.subject_entry_var, 
                                 font=('Arial', 11), width=30)
        subject_entry.pack(side=tk.LEFT, padx=10)
        
        ttk.Button(add_frame, text="Add Subject", 
                  command=self.add_subject).pack(side=tk.LEFT)
        
        # Subjects list frame
        list_frame = ttk.LabelFrame(tab, text="Current Subjects", padding=20)
        list_frame.pack(pady=20, padx=50, fill='both', expand=True)
        
        # Create listbox with scrollbar
        self.subjects_listbox = tk.Listbox(list_frame, font=('Arial', 11), 
                                          height=12, selectmode=tk.SINGLE)
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, 
                                 command=self.subjects_listbox.yview)
        self.subjects_listbox.configure(yscrollcommand=scrollbar.set)
        
        self.subjects_listbox.pack(side=tk.LEFT, fill='both', expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # List management buttons
        button_frame = ttk.Frame(list_frame)
        button_frame.pack(side=tk.BOTTOM, pady=10, fill='x')
        
        ttk.Button(button_frame, text="Remove Selected", 
                  command=self.remove_subject).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Clear All", 
                  command=self.clear_subjects).pack(side=tk.LEFT, padx=5)
    
    def create_classes_tab(self):
        """Create classes/standards management tab"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Classes")
        
        # Title
        title = ttk.Label(tab, text="Classes and Sections Management", 
                         font=('Arial', 16, 'bold'))
        title.pack(pady=20)
        
        # Notebook for standards and sections
        class_notebook = ttk.Notebook(tab)
        class_notebook.pack(pady=10, padx=50, fill='both', expand=True)
        
        # Standards tab
        standards_tab = ttk.Frame(class_notebook)
        class_notebook.add(standards_tab, text="Standards")
        
        # Add standard frame
        add_std_frame = ttk.Frame(standards_tab)
        add_std_frame.pack(pady=10, fill='x')
        
        ttk.Label(add_std_frame, text="Standard/Class Name:", 
                 font=('Arial', 11)).pack(side=tk.LEFT)
        self.std_entry_var = tk.StringVar()
        ttk.Entry(add_std_frame, textvariable=self.std_entry_var, 
                 font=('Arial', 11), width=20).pack(side=tk.LEFT, padx=10)
        
        ttk.Button(add_std_frame, text="Add Standard", 
                  command=self.add_standard).pack(side=tk.LEFT)
        
        # Standards list
        std_list_frame = ttk.LabelFrame(standards_tab, text="Standards", padding=10)
        std_list_frame.pack(pady=10, fill='both', expand=True)
        
        self.std_listbox = tk.Listbox(std_list_frame, font=('Arial', 11), 
                                     height=10, selectmode=tk.SINGLE)
        std_scrollbar = ttk.Scrollbar(std_list_frame, orient=tk.VERTICAL, 
                                     command=self.std_listbox.yview)
        self.std_listbox.configure(yscrollcommand=std_scrollbar.set)
        
        self.std_listbox.pack(side=tk.LEFT, fill='both', expand=True)
        std_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Sections tab
        sections_tab = ttk.Frame(class_notebook)
        class_notebook.add(sections_tab, text="Sections")
        
        # Standard selection for sections
        std_select_frame = ttk.Frame(sections_tab)
        std_select_frame.pack(pady=10, fill='x')
        
        ttk.Label(std_select_frame, text="Select Standard:", 
                 font=('Arial', 11)).pack(side=tk.LEFT)
        self.std_combo_var = tk.StringVar()
        self.std_combo = ttk.Combobox(std_select_frame, 
                                     textvariable=self.std_combo_var, 
                                     state='readonly', width=20)
        self.std_combo.pack(side=tk.LEFT, padx=10)
        
        # Add section frame
        add_sec_frame = ttk.Frame(sections_tab)
        add_sec_frame.pack(pady=10, fill='x')
        
        ttk.Label(add_sec_frame, text="Section Name:", 
                 font=('Arial', 11)).pack(side=tk.LEFT)
        self.sec_entry_var = tk.StringVar()
        ttk.Entry(add_sec_frame, textvariable=self.sec_entry_var, 
                 font=('Arial', 11), width=20).pack(side=tk.LEFT, padx=10)
        
        ttk.Button(add_sec_frame, text="Add Section", 
                  command=self.add_section).pack(side=tk.LEFT)
        
        # Sections list
        sec_list_frame = ttk.LabelFrame(sections_tab, text="Sections", padding=10)
        sec_list_frame.pack(pady=10, fill='both', expand=True)
        
        self.sec_tree = ttk.Treeview(sec_list_frame, columns=('Standard', 'Section'), 
                                    show='headings', height=10)
        self.sec_tree.heading('Standard', text='Standard')
        self.sec_tree.heading('Section', text='Section')
        self.sec_tree.column('Standard', width=150)
        self.sec_tree.column('Section', width=150)
        
        sec_scrollbar = ttk.Scrollbar(sec_list_frame, orient=tk.VERTICAL, 
                                     command=self.sec_tree.yview)
        self.sec_tree.configure(yscrollcommand=sec_scrollbar.set)
        
        self.sec_tree.pack(side=tk.LEFT, fill='both', expand=True)
        sec_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # List management buttons
        button_frame = ttk.Frame(sections_tab)
        button_frame.pack(pady=10, fill='x')
        
        ttk.Button(button_frame, text="Remove Selected", 
                  command=self.remove_section).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Clear All", 
                  command=self.clear_sections).pack(side=tk.LEFT, padx=5)
    
    def create_teachers_tab(self):
        """Create teachers management tab"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Teachers")
        
        # Title
        title = ttk.Label(tab, text="Teachers Management", 
                         font=('Arial', 16, 'bold'))
        title.pack(pady=20)
        
        # Add teacher frame
        add_frame = ttk.LabelFrame(tab, text="Add Teacher with Teaching Hours", padding=15)
        add_frame.pack(pady=10, padx=50, fill='x')
        
        # Teacher name row
        name_row = ttk.Frame(add_frame)
        name_row.pack(fill='x', pady=5)
        
        ttk.Label(name_row, text="Teacher Name:", 
                 font=('Arial', 11), width=15).pack(side=tk.LEFT)
        self.teacher_entry_var = tk.StringVar()
        teacher_entry = ttk.Entry(name_row, textvariable=self.teacher_entry_var, 
                                 font=('Arial', 11), width=25)
        teacher_entry.pack(side=tk.LEFT, padx=10)
        
        # Teaching hours row
        hours_row = ttk.Frame(add_frame)
        hours_row.pack(fill='x', pady=5)
        
        ttk.Label(hours_row, text="Min Hours/Day:", 
                 font=('Arial', 11), width=15).pack(side=tk.LEFT)
        self.teacher_min_hours_var = tk.IntVar(value=3)
        min_hours_entry = ttk.Entry(hours_row, textvariable=self.teacher_min_hours_var, 
                                   font=('Arial', 11), width=10)
        min_hours_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(hours_row, text="Max Hours/Day:", 
                 font=('Arial', 11)).pack(side=tk.LEFT, padx=(20, 5))
        self.teacher_max_hours_var = tk.IntVar(value=4)
        max_hours_entry = ttk.Entry(hours_row, textvariable=self.teacher_max_hours_var, 
                                   font=('Arial', 11), width=10)
        max_hours_entry.pack(side=tk.LEFT, padx=5)
        
        # Add button
        button_row = ttk.Frame(add_frame)
        button_row.pack(fill='x', pady=10)
        
        ttk.Button(button_row, text="Add Teacher with Hours", 
                  command=self.add_teacher_with_hours).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_row, text="Add Teacher Only", 
                  command=self.add_teacher).pack(side=tk.LEFT, padx=5)
        
        # Quick add multiple teachers frame
        quick_frame = ttk.LabelFrame(tab, text="Quick Add Multiple Teachers", padding=15)
        quick_frame.pack(pady=10, padx=50, fill='x')
        
        quick_input_row = ttk.Frame(quick_frame)
        quick_input_row.pack(fill='x', pady=5)
        
        ttk.Label(quick_input_row, text="Teachers (comma separated):", 
                 font=('Arial', 11), width=20).pack(side=tk.LEFT)
        self.quick_teachers_var = tk.StringVar()
        quick_entry = ttk.Entry(quick_input_row, textvariable=self.quick_teachers_var, 
                               font=('Arial', 11), width=40)
        quick_entry.pack(side=tk.LEFT, padx=10, fill='x', expand=True)
        
        quick_hours_row = ttk.Frame(quick_frame)
        quick_hours_row.pack(fill='x', pady=5)
        
        ttk.Label(quick_hours_row, text="Default Hours (Min-Max):", 
                 font=('Arial', 11), width=20).pack(side=tk.LEFT)
        self.quick_min_hours_var = tk.IntVar(value=3)
        quick_min_entry = ttk.Entry(quick_hours_row, textvariable=self.quick_min_hours_var, 
                                   font=('Arial', 11), width=8)
        quick_min_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(quick_hours_row, text="-").pack(side=tk.LEFT, padx=2)
        
        self.quick_max_hours_var = tk.IntVar(value=4)
        quick_max_entry = ttk.Entry(quick_hours_row, textvariable=self.quick_max_hours_var, 
                                   font=('Arial', 11), width=8)
        quick_max_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(quick_frame, text="Add Multiple Teachers", 
                  command=self.add_multiple_teachers).pack(pady=10)
        
        # Teachers list frame
        list_frame = ttk.LabelFrame(tab, text="Current Teachers", padding=20)
        list_frame.pack(pady=20, padx=50, fill='both', expand=True)
        
        # Create listbox with scrollbar
        self.teachers_listbox = tk.Listbox(list_frame, font=('Arial', 11), 
                                          height=12, selectmode=tk.SINGLE)
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, 
                                 command=self.teachers_listbox.yview)
        self.teachers_listbox.configure(yscrollcommand=scrollbar.set)
        
        self.teachers_listbox.pack(side=tk.LEFT, fill='both', expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # List management buttons
        button_frame = ttk.Frame(list_frame)
        button_frame.pack(side=tk.BOTTOM, pady=10, fill='x')
        
        ttk.Button(button_frame, text="Remove Selected", 
                  command=self.remove_teacher).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Clear All", 
                  command=self.clear_teachers).pack(side=tk.LEFT, padx=5)
    
    def create_assignments_tab(self):
        """Create teacher-subject and class-subject assignment tab"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Assignments")
        
        # Create notebook for assignments
        assign_notebook = ttk.Notebook(tab)
        assign_notebook.pack(pady=10, padx=10, fill='both', expand=True)
        
        # Teacher-Subject Assignment Tab
        ts_tab = ttk.Frame(assign_notebook)
        assign_notebook.add(ts_tab, text="Teacher-Subject")
        
        # Teacher selection
        teacher_frame = ttk.Frame(ts_tab)
        teacher_frame.pack(pady=10, fill='x')
        
        ttk.Label(teacher_frame, text="Select Teacher:", 
                 font=('Arial', 11)).pack(side=tk.LEFT)
        self.ts_teacher_var = tk.StringVar()
        self.ts_teacher_combo = ttk.Combobox(teacher_frame, 
                                           textvariable=self.ts_teacher_var, 
                                           state='readonly', width=25)
        self.ts_teacher_combo.pack(side=tk.LEFT, padx=10)
        self.ts_teacher_combo.bind('<<ComboboxSelected>>', self.on_teacher_selected)
        
        # Subject assignment frame
        assign_frame = ttk.LabelFrame(ts_tab, text="Assign Subjects to Teacher (Multiple selection allowed)", padding=20)
        assign_frame.pack(pady=10, padx=20, fill='both', expand=True)
        
        # Create a frame for subject checkboxes
        self.ts_subjects_frame = ttk.Frame(assign_frame)
        self.ts_subjects_frame.pack(fill='both', expand=True)
        
        # Initially show placeholder
        ttk.Label(self.ts_subjects_frame, 
                 text="Select a teacher to assign subjects",
                 font=('Arial', 10, 'italic')).pack(pady=20)
        
        # Dictionary to store subject checkboxes variables
        self.ts_subject_vars = {}
        
        # Assignment list
        list_frame = ttk.LabelFrame(ts_tab, text="Subject-Teacher Assignments (Subject → Teachers)", padding=10)
        list_frame.pack(pady=10, padx=20, fill='both', expand=True)
        
        self.ts_tree = ttk.Treeview(list_frame, columns=('Subject', 'Teachers'), 
                                   show='headings', height=8)
        self.ts_tree.heading('Subject', text='Subject')
        self.ts_tree.heading('Teachers', text='Teachers')
        self.ts_tree.column('Subject', width=150)
        self.ts_tree.column('Teachers', width=300)
        
        ts_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, 
                                    command=self.ts_tree.yview)
        self.ts_tree.configure(yscrollcommand=ts_scrollbar.set)
        
        self.ts_tree.pack(side=tk.LEFT, fill='both', expand=True)
        ts_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Assignment buttons
        button_frame = ttk.Frame(ts_tab)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="Save Teacher Subjects", 
                  command=self.assign_teacher_subject).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Remove Assignment", 
                  command=self.remove_teacher_subject).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Refresh Display", 
                  command=self.refresh_teacher_subjects).pack(side=tk.LEFT, padx=5)
        
        # Class-Subject Assignment Tab
        cs_tab = ttk.Frame(assign_notebook)
        assign_notebook.add(cs_tab, text="Class-Subject")
        
        # Class selection
        class_frame = ttk.Frame(cs_tab)
        class_frame.pack(pady=10, fill='x')
        
        ttk.Label(class_frame, text="Select Class:", 
                 font=('Arial', 11)).pack(side=tk.LEFT)
        self.cs_class_var = tk.StringVar()
        self.cs_class_combo = ttk.Combobox(class_frame, 
                                         textvariable=self.cs_class_var, 
                                         state='readonly', width=25)
        self.cs_class_combo.pack(side=tk.LEFT, padx=10)
        self.cs_class_combo.bind('<<ComboboxSelected>>', self.on_class_selected)
        
        # Subject assignment frame
        cs_assign_frame = ttk.LabelFrame(cs_tab, text="Assign Subjects to Class", padding=20)
        cs_assign_frame.pack(pady=10, padx=20, fill='both', expand=True)
        
        # Create a frame for subject checkboxes
        self.cs_subjects_frame = ttk.Frame(cs_assign_frame)
        self.cs_subjects_frame.pack(fill='both', expand=True)
        
        # Initially show placeholder
        ttk.Label(self.cs_subjects_frame, 
                 text="Select a class to assign subjects",
                 font=('Arial', 10, 'italic')).pack(pady=20)
        
        # Dictionary to store subject checkboxes variables
        self.cs_subject_vars = {}
        
        # Assignment list
        cs_list_frame = ttk.LabelFrame(cs_tab, text="Class-Subject Assignments", padding=10)
        cs_list_frame.pack(pady=10, padx=20, fill='both', expand=True)
        
        self.cs_tree = ttk.Treeview(cs_list_frame, columns=('Class', 'Subjects'), 
                                   show='headings', height=8)
        self.cs_tree.heading('Class', text='Class')
        self.cs_tree.heading('Subjects', text='Subjects')
        self.cs_tree.column('Class', width=150)
        self.cs_tree.column('Subjects', width=300)
        
        cs_scrollbar = ttk.Scrollbar(cs_list_frame, orient=tk.VERTICAL, 
                                    command=self.cs_tree.yview)
        self.cs_tree.configure(yscrollcommand=cs_scrollbar.set)
        
        self.cs_tree.pack(side=tk.LEFT, fill='both', expand=True)
        cs_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Assignment buttons
        cs_button_frame = ttk.Frame(cs_tab)
        cs_button_frame.pack(pady=10)
        
        ttk.Button(cs_button_frame, text="Save Class Subjects", 
                  command=self.assign_class_subjects).pack(side=tk.LEFT, padx=5)
        ttk.Button(cs_button_frame, text="Clear Class Subjects", 
                  command=self.clear_class_subjects).pack(side=tk.LEFT, padx=5)
        ttk.Button(cs_button_frame, text="Refresh Display", 
                  command=self.refresh_class_subjects).pack(side=tk.LEFT, padx=5)
    
    def create_generate_tab(self):
        """Create timetable generation tab"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Generate")
        
        # Title
        title = ttk.Label(tab, text="Generate Optimized Timetable", 
                         font=('Arial', 16, 'bold'))
        title.pack(pady=20)
        
        # Status frame
        status_frame = ttk.LabelFrame(tab, text="Generation Status", padding=20)
        status_frame.pack(pady=10, padx=50, fill='x')
        
        self.gen_status_text = scrolledtext.ScrolledText(status_frame, 
                                                        height=15, 
                                                        font=('Consolas', 9))
        self.gen_status_text.pack(fill='both', expand=True)
        
        # Options frame
        options_frame = ttk.LabelFrame(tab, text="Generation Options", padding=20)
        options_frame.pack(pady=10, padx=50, fill='x')
        
        # Algorithm selection
        ttk.Label(options_frame, text="Generation Algorithm:", 
                 font=('Arial', 11)).grid(row=0, column=0, sticky=tk.W, pady=10)
        self.algo_var = tk.StringVar(value="Smart Distribution")
        algo_combo = ttk.Combobox(options_frame, textvariable=self.algo_var, 
                                 values=["Smart Distribution", "Round Robin", "Random"], 
                                 state='readonly', width=20)
        algo_combo.grid(row=0, column=1, pady=10, padx=10, sticky=tk.W)
        
        # Conflict avoidance
        self.avoid_conflicts_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Avoid Teacher Conflicts (Recommended)", 
                       variable=self.avoid_conflicts_var).grid(row=1, column=0, 
                                                             columnspan=2, 
                                                             sticky=tk.W, pady=5)
        
        self.balance_load_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Balance Teacher Workload", 
                       variable=self.balance_load_var).grid(row=2, column=0, 
                                                          columnspan=2, 
                                                          sticky=tk.W, pady=5)
        
        self.avoid_same_subject_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Avoid Same Subject in Parallel Sections", 
                       variable=self.avoid_same_subject_var).grid(row=3, column=0, 
                                                                columnspan=2, 
                                                                sticky=tk.W, pady=5)
        
        # Buttons frame
        button_frame = ttk.Frame(tab)
        button_frame.pack(pady=30)
        
        ttk.Button(button_frame, text="🔄 Validate Data", 
                  command=self.validate_data, 
                  style='Accent.TButton').pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="🚀 Generate Timetable", 
                  command=self.generate_timetable, 
                  style='Accent.TButton').pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="🔄 Regenerate All", 
                  command=self.regenerate_all).pack(side=tk.LEFT, padx=10)
    
    def create_view_tab(self):
        """Create timetable viewing tab"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="View")
        
        # Title
        title = ttk.Label(tab, text="View Generated Timetables", 
                         font=('Arial', 16, 'bold'))
        title.pack(pady=20)
        
        # Control frame
        control_frame = ttk.Frame(tab)
        control_frame.pack(pady=10, padx=50, fill='x')
        
        # Class selection
        ttk.Label(control_frame, text="Select Class/Section:", 
                 font=('Arial', 11)).pack(side=tk.LEFT)
        self.view_class_var = tk.StringVar()
        self.view_class_combo = ttk.Combobox(control_frame, 
                                           textvariable=self.view_class_var, 
                                           state='readonly', width=25)
        self.view_class_combo.pack(side=tk.LEFT, padx=10)
        
        # Day selection
        ttk.Label(control_frame, text="Day:", 
                 font=('Arial', 11)).pack(side=tk.LEFT, padx=(20, 5))
        self.view_day_var = tk.StringVar(value="All")
        day_combo = ttk.Combobox(control_frame, textvariable=self.view_day_var, 
                                values=["All"] + self.days, 
                                state='readonly', width=15)
        day_combo.pack(side=tk.LEFT, padx=5)
        
        # View buttons
        ttk.Button(control_frame, text="📋 View Text", 
                  command=self.view_text_timetable).pack(side=tk.LEFT, padx=20)
        ttk.Button(control_frame, text="📊 View Teacher Schedule", 
                  command=self.view_teacher_schedule).pack(side=tk.LEFT, padx=5)
        
        # Display frame
        display_frame = ttk.Frame(tab)
        display_frame.pack(pady=10, padx=10, fill='both', expand=True)
        
        # Create notebook for different views
        view_notebook = ttk.Notebook(display_frame)
        view_notebook.pack(fill='both', expand=True)
        
        # Text view tab
        text_view = ttk.Frame(view_notebook)
        self.view_text = scrolledtext.ScrolledText(text_view, 
                                                  font=('Consolas', 10), 
                                                  wrap=tk.WORD)
        self.view_text.pack(fill='both', expand=True)
        view_notebook.add(text_view, text="Timetable View")
        
        # Teacher schedule tab
        teacher_view = ttk.Frame(view_notebook)
        self.teacher_text = scrolledtext.ScrolledText(teacher_view, 
                                                     font=('Consolas', 10), 
                                                     wrap=tk.WORD)
        self.teacher_text.pack(fill='both', expand=True)
        view_notebook.add(teacher_view, text="Teacher Schedule")
    
    def create_export_tab(self):
        """Create export options tab"""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Export")
        
        # Title
        title = ttk.Label(tab, text="Export Timetable", 
                         font=('Arial', 16, 'bold'))
        title.pack(pady=20)
        
        # Export options frame
        options_frame = ttk.LabelFrame(tab, text="Export Options", padding=30)
        options_frame.pack(pady=10, padx=50, fill='both', expand=True)
        
        # Excel export
        excel_frame = ttk.Frame(options_frame)
        excel_frame.pack(pady=10, fill='x')
        
        ttk.Label(excel_frame, text="Export to Excel:", 
                 font=('Arial', 11, 'bold')).pack(side=tk.LEFT)
        ttk.Button(excel_frame, text="📊 Export All to Excel", 
                  command=self.export_excel).pack(side=tk.LEFT, padx=20)
        ttk.Button(excel_frame, text="📁 Export Selected...", 
                  command=self.export_excel_selected).pack(side=tk.LEFT, padx=5)
        
        # JSON export
        json_frame = ttk.Frame(options_frame)
        json_frame.pack(pady=10, fill='x')
        
        ttk.Label(json_frame, text="Save Configuration:", 
                 font=('Arial', 11, 'bold')).pack(side=tk.LEFT)
        ttk.Button(json_frame, text="💾 Save as JSON", 
                  command=self.save_configuration).pack(side=tk.LEFT, padx=20)
        ttk.Button(json_frame, text="📂 Load from JSON", 
                  command=self.load_configuration).pack(side=tk.LEFT, padx=5)
        
        # Summary export
        summary_frame = ttk.Frame(options_frame)
        summary_frame.pack(pady=10, fill='x')
        
        ttk.Label(summary_frame, text="Export Summary:", 
                 font=('Arial', 11, 'bold')).pack(side=tk.LEFT)
        ttk.Button(summary_frame, text="📝 Export Summary Report", 
                  command=self.export_summary).pack(side=tk.LEFT, padx=20)
        
        # Status display
        status_frame = ttk.LabelFrame(tab, text="Export Status", padding=20)
        status_frame.pack(pady=20, padx=50, fill='x')
        
        self.export_status_text = scrolledtext.ScrolledText(status_frame, 
                                                          height=6, 
                                                          font=('Consolas', 10))
        self.export_status_text.pack(fill='both', expand=True)
    
    # ========== HELPER METHODS ==========
    
    def update_status(self, message):
        """Update status bar"""
        self.status_var.set(message)
        self.root.update()
    
    def log_generation(self, message):
        """Log message to generation status"""
        self.gen_status_text.insert(tk.END, message + "\n")
        self.gen_status_text.see(tk.END)
        self.root.update()
    
    def log_export(self, message):
        """Log message to export status"""
        self.export_status_text.insert(tk.END, message + "\n")
        self.export_status_text.see(tk.END)
        self.root.update()
    
    # ========== SCHOOL TIMING METHODS ==========
    
    def save_school_timing(self):
        """Save school timing configuration"""
        try:
            start_time = self.start_time_var.get()
            end_time = self.end_time_var.get()
            
            # Validate time format
            datetime.strptime(start_time, "%H:%M")
            datetime.strptime(end_time, "%H:%M")
            
            # Get selected days
            selected_days = [day for day, var in self.day_vars.items() if var.get()]
            if not selected_days:
                messagebox.showerror("Error", "Please select at least one working day")
                return
            
            self.school_timing = {
                'start': start_time,
                'end': end_time,
                'duration': self.calculate_duration(start_time, end_time)
            }
            
            self.days = selected_days
            
            # Update preview
            preview = f"School Timing: {start_time} to {end_time} ({self.school_timing['duration']} minutes)\n"
            preview += f"Working Days: {', '.join(selected_days)}"
            self.timing_preview.config(text=preview)
            
            self.update_status("School timing saved successfully")
            messagebox.showinfo("Success", "School timing configuration saved!")
            
        except ValueError:
            messagebox.showerror("Error", "Invalid time format. Please use HH:MM format")
    
    def calculate_duration(self, start_time, end_time):
        """Calculate duration between two times"""
        start = datetime.strptime(start_time, "%H:%M")
        end = datetime.strptime(end_time, "%H:%M")
        duration = (end - start).seconds // 60
        return duration
    
    # ========== PERIODS METHODS ==========
    
    def add_recess(self):
        """Add a recess period"""
        name = self.recess_name_var.get()
        start_time = self.recess_time_var.get()
        duration = self.recess_duration_var.get()
        
        if not name or not start_time:
            messagebox.showerror("Error", "Please enter recess name and start time")
            return
        
        try:
            datetime.strptime(start_time, "%H:%M")
            self.recess_tree.insert('', 'end', values=(name, start_time, f"{duration} min"))
            
            # Clear inputs to correct defaults
            self.recess_name_var.set("Recess")
            self.recess_time_var.set("11:30")
            self.recess_duration_var.set(30)
            
        except ValueError:
            messagebox.showerror("Error", "Invalid time format. Please use HH:MM format")
    
    def remove_recess(self):
        """Remove selected recess"""
        selection = self.recess_tree.selection()
        if selection:
            self.recess_tree.delete(selection)
    
    def save_period_settings(self):
        """Save period configuration"""
        try:
            # Get recess times
            recess_times = []
            for item in self.recess_tree.get_children():
                values = self.recess_tree.item(item)['values']
                recess_times.append({
                    'name': values[0],
                    'start': values[1],
                    'duration': int(values[2].split()[0])
                })
            
            # Update config
            self.config.update({
                'period_duration': self.period_duration_var.get(),
                'period_overtime': self.period_overtime_var.get(),
                'recess_times': recess_times,
                'teaching_hours': self.teaching_hours_var.get(),
                'school_timing': self.school_timing,
                'days': self.days
            })
            
            self.update_status("Period settings saved successfully")
            messagebox.showinfo("Success", "Period configuration saved!")
            
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid input: {str(e)}")
    
    # ========== SUBJECTS METHODS ==========
    
    def add_subject(self):
        """Add a new subject"""
        subject = self.subject_entry_var.get().strip()
        if subject and subject not in self.subjects:
            self.subjects.append(subject)
            self.subjects_listbox.insert(tk.END, subject)
            self.subject_entry_var.set("")
            
            # Update assignment displays
            self.update_assignment_comboboxes()
            
            self.update_status(f"Subject '{subject}' added")
    
    def remove_subject(self):
        """Remove selected subject"""
        selection = self.subjects_listbox.curselection()
        if selection:
            subject = self.subjects_listbox.get(selection[0])
            self.subjects.remove(subject)
            self.subjects_listbox.delete(selection[0])
            
            # Remove from assignments
            if subject in self.subject_teachers:
                del self.subject_teachers[subject]
            
            # Remove from standard subjects
            for standard in self.standard_subjects:
                if subject in self.standard_subjects[standard]:
                    self.standard_subjects[standard].remove(subject)
            
            # Update assignment displays
            self.update_assignment_comboboxes()
            
            self.update_status(f"Subject '{subject}' removed")
    
    def clear_subjects(self):
        """Clear all subjects"""
        if messagebox.askyesno("Confirm", "Clear all subjects? This will also clear all assignments."):
            self.subjects.clear()
            self.subject_teachers.clear()
            self.subjects_listbox.delete(0, tk.END)
            
            # Clear standard subjects
            for standard in self.standard_subjects:
                self.standard_subjects[standard] = []
            
            # Update assignment displays
            self.update_assignment_comboboxes()
            
            self.update_status("All subjects cleared")
    
    # ========== CLASSES METHODS ==========
    
    def add_standard(self):
        """Add a new standard"""
        standard = self.std_entry_var.get().strip()
        if standard and standard not in self.standards:
            self.standards.append(standard)
            self.std_listbox.insert(tk.END, standard)
            self.std_entry_var.set("")
            
            # Update comboboxes
            self.std_combo['values'] = self.standards
            self.cs_class_combo['values'] = self.standards
            
            # Initialize empty subject list for this standard
            self.standard_subjects[standard] = []
            
            self.update_status(f"Standard '{standard}' added")
    
    def add_section(self):
        """Add a new section"""
        standard = self.std_combo_var.get()
        section = self.sec_entry_var.get().strip()
        
        if not standard:
            messagebox.showerror("Error", "Please select a standard first")
            return
        
        if not section:
            messagebox.showerror("Error", "Please enter section name")
            return
        
        section_full = f"{standard}-{section}"
        if section_full not in self.sections:
            self.sections.append(section_full)
            self.sec_tree.insert('', 'end', values=(standard, section))
            self.sec_entry_var.set("")
            
            # Update view combobox
            self.view_class_combo['values'] = self.sections
            
            self.update_status(f"Section '{section_full}' added")
    
    def remove_section(self):
        """Remove selected section"""
        selection = self.sec_tree.selection()
        if selection:
            values = self.sec_tree.item(selection[0])['values']
            section_full = f"{values[0]}-{values[1]}"
            
            if section_full in self.sections:
                self.sections.remove(section_full)
                self.sec_tree.delete(selection[0])
                
                # Remove from timetable data if exists
                if section_full in self.timetable_data:
                    del self.timetable_data[section_full]
                
                # Update view combobox
                self.view_class_combo['values'] = self.sections
                
                self.update_status(f"Section '{section_full}' removed")
    
    def clear_sections(self):
        """Clear all sections"""
        if messagebox.askyesno("Confirm", "Clear all sections?"):
            self.sections.clear()
            for item in self.sec_tree.get_children():
                self.sec_tree.delete(item)
            
            # Clear timetable data
            self.timetable_data.clear()
            
            self.update_status("All sections cleared")
    
    # ========== TEACHERS METHODS ==========
    
    def add_teacher(self):
        """Add a new teacher"""
        teacher = self.teacher_entry_var.get().strip()
        if teacher and teacher not in self.teachers:
            self.teachers.append(teacher)
            self.teachers_listbox.insert(tk.END, teacher)
            self.teacher_entry_var.set("")
            
            # Update combobox
            self.ts_teacher_combo['values'] = self.teachers
            
            self.update_status(f"Teacher '{teacher}' added")
    
    def add_teacher_with_hours(self):
        """Add a new teacher with teaching hours"""
        teacher = self.teacher_entry_var.get().strip()
        if not teacher:
            messagebox.showerror("Error", "Please enter teacher name")
            return
        
        if teacher in self.teachers:
            messagebox.showerror("Error", f"Teacher '{teacher}' already exists")
            return
        
        min_hours = self.teacher_min_hours_var.get()
        max_hours = self.teacher_max_hours_var.get()
        
        if min_hours > max_hours:
            messagebox.showerror("Error", "Minimum hours cannot be greater than maximum hours")
            return
        
        # Add teacher
        self.teachers.append(teacher)
        self.teachers_listbox.insert(tk.END, teacher)
        
        # Store teacher hours
        self.teacher_hours[teacher] = {
            'min_hours': min_hours,
            'max_hours': max_hours
        }
        
        # Clear inputs
        self.teacher_entry_var.set("")
        
        # Update comboboxes and trees
        self.update_assignment_comboboxes()
        
        self.update_status(f"Teacher '{teacher}' added with {min_hours}-{max_hours} hours/day")
        messagebox.showinfo("Success", f"Teacher '{teacher}' added successfully!")
    
    def add_multiple_teachers(self):
        """Add multiple teachers with default hours"""
        teachers_text = self.quick_teachers_var.get().strip()
        if not teachers_text:
            messagebox.showerror("Error", "Please enter teacher names")
            return
        
        min_hours = self.quick_min_hours_var.get()
        max_hours = self.quick_max_hours_var.get()
        
        if min_hours > max_hours:
            messagebox.showerror("Error", "Minimum hours cannot be greater than maximum hours")
            return
        
        # Split by comma and clean names
        teacher_names = [name.strip() for name in teachers_text.split(',')]
        added_count = 0
        
        for teacher in teacher_names:
            if teacher and teacher not in self.teachers:
                self.teachers.append(teacher)
                self.teachers_listbox.insert(tk.END, teacher)
                
                # Store teacher hours
                self.teacher_hours[teacher] = {
                    'min_hours': min_hours,
                    'max_hours': max_hours
                }
                added_count += 1
        
        # Clear inputs
        self.quick_teachers_var.set("")
        
        # Update comboboxes and trees
        self.update_assignment_comboboxes()
        
        if added_count > 0:
            self.update_status(f"Added {added_count} teachers with {min_hours}-{max_hours} hours/day")
            messagebox.showinfo("Success", f"Successfully added {added_count} teachers!")
        else:
            messagebox.showwarning("Warning", "No new teachers were added")
    
    def remove_teacher(self):
        """Remove selected teacher"""
        selection = self.teachers_listbox.curselection()
        if selection:
            teacher = self.teachers_listbox.get(selection[0])
            self.teachers.remove(teacher)
            self.teachers_listbox.delete(selection[0])
            
            # Remove teacher from all subject assignments
            for subject in list(self.subject_teachers.keys()):
                if subject in self.subject_teachers:
                    if teacher in self.subject_teachers[subject]:
                        self.subject_teachers[subject].remove(teacher)
                        # If no teachers left for this subject, remove the entry
                        if not self.subject_teachers[subject]:
                            del self.subject_teachers[subject]
            
            # Update combobox
            self.ts_teacher_combo['values'] = self.teachers
            
            # Update assignment display
            self.update_assignment_comboboxes()
            
            self.update_status(f"Teacher '{teacher}' removed")
    
    def clear_teachers(self):
        """Clear all teachers"""
        if messagebox.askyesno("Confirm", "Clear all teachers? This will also clear all teacher assignments."):
            self.teachers.clear()
            self.subject_teachers.clear()
            self.teachers_listbox.delete(0, tk.END)
            
            # Update assignment displays
            self.update_assignment_comboboxes()
            
            self.update_status("All teachers cleared")
    
    # ========== TEACHER HOURS METHODS ==========
    
    def set_teacher_hours(self):
        """Set teaching hours for selected teacher"""
        # For now, show info message since UI components are not fully set up
        messagebox.showinfo("Info", "Please use the Teachers tab to set teaching hours when adding teachers.\n\nYou can also add teachers with hours directly using:\n- 'Add Teacher with Hours' button\n- 'Quick Add Multiple Teachers' option")
    
    def clear_teacher_hours(self):
        """Clear all teacher hours"""
        if messagebox.askyesno("Confirm", "Clear all teacher hours?"):
            self.teacher_hours.clear()
            self.update_teacher_hours_tree()
            self.update_status("All teacher hours cleared")
    
    def validate_teacher_hours(self):
        """Validate teacher hours against current assignments"""
        if not self.teacher_hours:
            messagebox.showwarning("Warning", "No teacher hours configured. Please set teaching hours first.")
            return
        
        if not self.timetable_data:
            messagebox.showwarning("Warning", "No timetable generated. Please generate timetable first.")
            return
        
        # Calculate current daily teaching hours for each teacher
        current_daily_hours = {}
        for teacher in self.teachers:
            daily_hours = {}
            
            for section in self.timetable_data:
                for day in self.timetable_data[section]:
                    if day not in daily_hours:
                        daily_hours[day] = 0
                    
                    for period in self.timetable_data[section][day]:
                        if period['teacher'] == teacher and period['subject'] not in ['Recess', 'Free Period']:
                            # Convert periods to hours (35 minutes per period)
                            daily_hours[day] += (35 / 60)
            
            # Get average daily hours
            if daily_hours:
                avg_daily_hours = sum(daily_hours.values()) / len(daily_hours)
                current_daily_hours[teacher] = avg_daily_hours
            else:
                current_daily_hours[teacher] = 0
        
        # Check for teachers with insufficient hours
        insufficient_teachers = []
        for teacher, daily_hours in current_daily_hours.items():
            if teacher in self.teacher_hours:
                min_required = self.teacher_hours[teacher]['min_hours']
                if daily_hours < min_required:
                    insufficient_teachers.append((teacher, daily_hours, min_required))
        
        if insufficient_teachers:
            # Create detailed message
            msg = "The following teachers have insufficient daily teaching hours:\n\n"
            for teacher, current, required in insufficient_teachers:
                msg += f"• {teacher}: {current:.1f} hours/day (required: {required} hours/day)\n"
            msg += f"\nPlease increase teaching hours for these teachers."
            
            messagebox.showwarning("Insufficient Teaching Hours", msg)
            self.update_status(f"Found {len(insufficient_teachers)} teachers with insufficient daily hours")
        else:
            messagebox.showinfo("Validation Success", "All teachers meet minimum daily teaching hour requirements!")
            self.update_status("All teacher daily hours validated successfully")
    
    def update_teacher_hours_tree(self):
        """Update teacher hours tree display"""
        # For now, skip tree update since UI component is not fully set up
        # Tree display will be added when Periods tab UI is completed
        pass
    
    # ========== ASSIGNMENT METHODS ==========
    
    def update_assignment_comboboxes(self):
        """Update combobox values in assignments tab"""
        self.ts_teacher_combo['values'] = self.teachers
        self.cs_class_combo['values'] = self.standards
        
        # Refresh teacher and class subject displays if they have selections
        if self.ts_teacher_var.get():
            self.on_teacher_selected()
        if self.cs_class_var.get():
            self.on_class_selected()
        
        # Update assignment trees
        self.update_assignment_trees()
        
        # Update teacher hours tree
        self.update_teacher_hours_tree()
    
    def update_assignment_trees(self):
        """Update the assignment trees with current data"""
        # Update teacher-subject tree
        for item in self.ts_tree.get_children():
            self.ts_tree.delete(item)
        
        # Display subjects with their teachers
        for subject in self.subjects:
            teachers = self.subject_teachers.get(subject, [])
            if teachers:
                self.ts_tree.insert('', 'end', values=(subject, ', '.join(teachers)))
        
        # Update class-subject tree
        for item in self.cs_tree.get_children():
            self.cs_tree.delete(item)
        
        for class_name, subjects in self.standard_subjects.items():
            if subjects:
                self.cs_tree.insert('', 'end', values=(class_name, ', '.join(subjects)))
    
    def on_teacher_selected(self, event=None):
        """When teacher is selected in teacher-subject assignment"""
        teacher = self.ts_teacher_var.get()
        
        # Clear the subjects frame
        for widget in self.ts_subjects_frame.winfo_children():
            widget.destroy()
        
        self.ts_subject_vars.clear()
        
        if not teacher:
            # Show placeholder
            ttk.Label(self.ts_subjects_frame, 
                     text="Select a teacher to assign subjects",
                     font=('Arial', 10, 'italic')).pack(pady=20)
            return
        
        if not self.subjects:
            # Show message if no subjects
            ttk.Label(self.ts_subjects_frame, 
                     text="No subjects available. Please add subjects first.",
                     font=('Arial', 10, 'italic')).pack(pady=20)
            return
        
        # Create a canvas and scrollbar for subjects
        canvas = tk.Canvas(self.ts_subjects_frame)
        scrollbar = ttk.Scrollbar(self.ts_subjects_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Create checkboxes for each subject
        for subject in self.subjects:
            # Check if this teacher is already assigned to this subject
            teachers_for_subject = self.subject_teachers.get(subject, [])
            is_assigned = teacher in teachers_for_subject
            
            var = tk.BooleanVar(value=is_assigned)
            chk = ttk.Checkbutton(scrollable_frame, text=subject, variable=var)
            chk.pack(anchor=tk.W, padx=10, pady=2)
            self.ts_subject_vars[subject] = var
        
        canvas.pack(side=tk.LEFT, fill='both', expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.update_status(f"Loaded subjects for {teacher}")
    
    def on_class_selected(self, event=None):
        """When class is selected in class-subject assignment"""
        class_name = self.cs_class_var.get()
        
        # Clear the subjects frame
        for widget in self.cs_subjects_frame.winfo_children():
            widget.destroy()
        
        self.cs_subject_vars.clear()
        
        if not class_name:
            # Show placeholder
            ttk.Label(self.cs_subjects_frame, 
                     text="Select a class to assign subjects",
                     font=('Arial', 10, 'italic')).pack(pady=20)
            return
        
        if not self.subjects:
            # Show message if no subjects
            ttk.Label(self.cs_subjects_frame, 
                     text="No subjects available. Please add subjects first.",
                     font=('Arial', 10, 'italic')).pack(pady=20)
            return
        
        # Create a canvas and scrollbar for subjects
        canvas = tk.Canvas(self.cs_subjects_frame)
        scrollbar = ttk.Scrollbar(self.cs_subjects_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Get subjects already assigned to this class
        assigned_subjects = self.standard_subjects.get(class_name, [])
        
        # Create checkboxes for each subject
        for subject in self.subjects:
            is_assigned = subject in assigned_subjects
            var = tk.BooleanVar(value=is_assigned)
            chk = ttk.Checkbutton(scrollable_frame, text=subject, variable=var)
            chk.pack(anchor=tk.W, padx=10, pady=2)
            self.cs_subject_vars[subject] = var
        
        canvas.pack(side=tk.LEFT, fill='both', expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.update_status(f"Loaded subjects for {class_name}")
    
    def refresh_teacher_subjects(self):
        """Refresh teacher-subject assignments display"""
        teacher = self.ts_teacher_var.get()
        if teacher:
            self.on_teacher_selected()
            self.update_status("Teacher subjects refreshed")
    
    def refresh_class_subjects(self):
        """Refresh class-subject assignments display"""
        class_name = self.cs_class_var.get()
        if class_name:
            self.on_class_selected()
            self.update_status("Class subjects refreshed")
    
    def assign_teacher_subject(self):
        """Assign selected subjects to teacher (multiple teachers can teach same subject)"""
        teacher = self.ts_teacher_var.get()
        if not teacher:
            messagebox.showerror("Error", "Please select a teacher")
            return
        
        if not self.ts_subject_vars:
            messagebox.showerror("Error", "No subjects available. Please add subjects first.")
            return
        
        # Get selected subjects
        selected_subjects = [subject for subject, var in self.ts_subject_vars.items() 
                           if var.get()]
        
        if not selected_subjects:
            messagebox.showwarning("Warning", "No subjects selected for assignment")
            return
        
        # Update assignments - multiple teachers can teach same subject
        for subject in selected_subjects:
            if subject not in self.subject_teachers:
                self.subject_teachers[subject] = []
            
            if teacher not in self.subject_teachers[subject]:
                self.subject_teachers[subject].append(teacher)
        
        # Remove teacher from subjects that are not selected
        for subject in self.subjects:
            if subject not in selected_subjects:
                if subject in self.subject_teachers and teacher in self.subject_teachers[subject]:
                    self.subject_teachers[subject].remove(teacher)
                    # If no teachers left for this subject, remove the subject entry
                    if not self.subject_teachers[subject]:
                        del self.subject_teachers[subject]
        
        # Update assignment trees
        self.update_assignment_trees()
        
        self.update_status(f"Updated subject assignments for {teacher}")
        messagebox.showinfo("Success", f"Updated assignments for {teacher}: {len(selected_subjects)} subjects")
    
    def remove_teacher_subject(self):
        """Remove entire subject assignment (whole cell/row)"""
        selection = self.ts_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a subject assignment to remove")
            return
        
        values = self.ts_tree.item(selection[0])['values']
        subject = values[0]
        teachers_str = values[1]
        
        # Confirm removal
        result = messagebox.askyesno("Confirm Removal", 
                                    f"Remove entire assignment for '{subject}'?\n\n"
                                    f"This will remove all teachers: {teachers_str}")
        
        if result:
            # Remove entire subject assignment
            if subject in self.subject_teachers:
                del self.subject_teachers[subject]
                self.update_assignment_trees()
                self.update_status(f"Removed entire assignment for {subject}")
                messagebox.showinfo("Success", f"Removed assignment for {subject}")
            else:
                messagebox.showerror("Error", f"Subject {subject} not found in assignments")
    
    def assign_class_subjects(self):
        """Assign selected subjects to class"""
        class_name = self.cs_class_var.get()
        if not class_name:
            messagebox.showerror("Error", "Please select a class")
            return
        
        if not self.cs_subject_vars:
            messagebox.showerror("Error", "No subjects available. Please add subjects first.")
            return
        
        # Get selected subjects
        selected_subjects = [subject for subject, var in self.cs_subject_vars.items() 
                           if var.get()]
        
        # Update assignments
        self.standard_subjects[class_name] = selected_subjects
        
        # Update assignment trees
        self.update_assignment_trees()
        
        self.update_status(f"Assigned {len(selected_subjects)} subjects to {class_name}")
        messagebox.showinfo("Success", f"Assigned {len(selected_subjects)} subjects to {class_name}")
    
    def clear_class_subjects(self):
        """Clear class-subject assignments"""
        class_name = self.cs_class_var.get()
        if class_name and class_name in self.standard_subjects:
            self.standard_subjects[class_name] = []
            
            # Update assignment trees
            self.update_assignment_trees()
            
            # Refresh the checkboxes
            self.on_class_selected()
            
            self.update_status(f"Subjects cleared for {class_name}")
    
    # ========== GENERATION METHODS ==========
    
    def validate_data(self):
        """Validate all input data before generation"""
        errors = []
        
        self.log_generation("=" * 60)
        self.log_generation("VALIDATING DATA...")
        
        # Check school timing
        if not self.school_timing:
            errors.append("School timing not configured")
        
        # Check periods
        if 'period_duration' not in self.config:
            errors.append("Period settings not configured")
        
        # Check subjects
        if not self.subjects:
            errors.append("No subjects added")
        
        # Check classes
        if not self.standards:
            errors.append("No standards/classes added")
        
        if not self.sections:
            errors.append("No sections added")
        
        # Check teachers
        if not self.teachers:
            errors.append("No teachers added")
        
        # Check teacher-subject assignments
        if not self.subject_teachers:
            errors.append("No teacher-subject assignments")
        else:
            # Check if all subjects have at least one teacher
            unassigned_subjects = [s for s in self.subjects if s not in self.subject_teachers]
            if unassigned_subjects:
                errors.append(f"Subjects without teacher assignment: {', '.join(unassigned_subjects)}")
        
        # Check class-subject assignments
        if not self.standard_subjects:
            errors.append("No class-subject assignments")
        else:
            for standard in self.standards:
                if standard not in self.standard_subjects or not self.standard_subjects[standard]:
                    errors.append(f"Standard {standard} has no subjects assigned")
        
        if errors:
            self.log_generation("✗ VALIDATION FAILED:")
            for error in errors:
                self.log_generation(f"  - {error}")
            messagebox.showerror("Validation Failed", 
                               "Please fix the following errors:\n\n• " + "\n• ".join(errors))
            return False
        else:
            self.log_generation("✓ ALL VALIDATIONS PASSED")
            
            # Show summary
            self.log_generation("\n--- CONFIGURATION SUMMARY ---")
            self.log_generation(f"Subjects: {len(self.subjects)}")
            self.log_generation(f"Teachers: {len(self.teachers)}")
            self.log_generation(f"Standards: {len(self.standards)}")
            self.log_generation(f"Sections: {len(self.sections)}")
            
            # Teacher-subject summary
            self.log_generation("\nSubject-Teacher Assignments (Multiple teachers per subject):")
            for subject, teachers in self.subject_teachers.items():
                self.log_generation(f"  {subject}: {len(teachers)} teachers ({', '.join(teachers)})")
        
            messagebox.showinfo("Success", "All data validated successfully!")
            return True
    
    def generate_time_slots(self, day):
        """Generate all time slots for a day based on school timing and recess configuration"""
        time_slots = []
        
        # Get configuration
        start_time = datetime.strptime(self.school_timing['start'], "%H:%M")
        end_time = datetime.strptime(self.school_timing['end'], "%H:%M")
        period_duration = self.config.get('period_duration', 35)
        period_overtime = self.config.get('period_overtime', 0)
        
        # Get recess times
        recess_times = self.config.get('recess_times', [])
        self.log_generation(f"DEBUG: School timing {start_time.strftime('%H:%M')} to {end_time.strftime('%H:%M')}")
        self.log_generation(f"DEBUG: Recess times found: {recess_times}")
        
        current_time = start_time
        period_count = 0
        
        while current_time < end_time:
            time_str = current_time.strftime("%H:%M")
            
            # Check if this time slot is a recess
            is_recess = False
            for recess in recess_times:
                recess_start = datetime.strptime(recess['start'], "%H:%M")
                if current_time.time() == recess_start.time():
                    # Add recess slot
                    time_slots.append({
                        'start': time_str,
                        'duration': recess['duration'],
                        'is_recess': True,
                        'name': recess['name']
                    })
                    self.log_generation(f"DEBUG: Added recess at {time_str} for {recess['duration']}min")
                    current_time += timedelta(minutes=recess['duration'])
                    is_recess = True
                    break
            
            if is_recess:
                continue
            
            # Check if we have enough time for a full period
            time_remaining = (end_time - current_time).total_seconds() / 60
            if time_remaining < period_duration:
                self.log_generation(f"DEBUG: Not enough time for full period at {time_str} (only {time_remaining:.0f}min left)")
                break
            
            # All slots are teaching periods (no free periods)
            slot_duration = period_duration + period_overtime
            
            # Add time slot
            time_slots.append({
                'start': time_str,
                'duration': slot_duration,
                'is_recess': False,
                'name': f'Period {period_count + 1}'
            })
            
            # Move to next time slot
            current_time += timedelta(minutes=slot_duration)
            period_count += 1
        
        self.log_generation(f"DEBUG: Generated {len(time_slots)} time slots for {day}")
        self.log_generation(f"DEBUG: Final time reached: {current_time.strftime('%H:%M')}")
        return time_slots
    
    def generate_timetable(self):
        """Generate optimized conflict-free timetable with multiple teachers per subject"""
        if not self.validate_data():
            return
        
        self.log_generation("=" * 60)
        self.log_generation("GENERATING OPTIMIZED TIMETABLE...")
        self.log_generation("Features: Multiple teachers per subject, No conflicts, Balanced workload")
        
        try:
            # Reset timetable data
            self.timetable_data = {}
            
            # Create teacher schedule tracker
            teacher_schedule = {}
            for teacher in self.teachers:
                teacher_schedule[teacher] = {}
                for day in self.days:
                    teacher_schedule[teacher][day] = []
            
            # Group sections by standard
            sections_by_standard = {}
            for section in self.sections:
                standard = section.split('-')[0]
                if standard not in sections_by_standard:
                    sections_by_standard[standard] = []
                sections_by_standard[standard].append(section)
            
            # Generate timetable for each standard
            for standard, sections in sections_by_standard.items():
                self.log_generation(f"\nGenerating for {standard} ({len(sections)} sections)...")
                
                # Get subjects for this standard
                standard_subjects = self.standard_subjects.get(standard, [])
                if not standard_subjects:
                    self.log_generation(f"  ⚠ No subjects assigned for {standard}")
                    continue
                
                # Generate for each day
                for day in self.days:
                    self.log_generation(f"  Scheduling {day}...")
                    
                    # Initialize schedules for all sections
                    day_schedules = {section: [] for section in sections}
                    periods_scheduled = {section: 0 for section in sections}
                    
                    # Generate all time slots for the day first
                    time_slots = self.generate_time_slots(day)
                    
                    period_num = 1
                    
                    for time_slot_info in time_slots:
                        current_time = datetime.strptime(time_slot_info['start'], "%H:%M")
                        time_slot = time_slot_info['start']
                        slot_duration = time_slot_info['duration']
                        
                        # Skip if this is a recess slot
                        if time_slot_info.get('is_recess', False):
                            for section in sections:
                                day_schedules[section].append({
                                    'period': 'R',
                                    'subject': time_slot_info['name'],
                                    'teacher': '',
                                    'time': time_slot,
                                    'duration': slot_duration
                                })
                            continue
                        
                        # Track assigned subjects and teachers for this time slot
                        assigned_subjects = []
                        assigned_teachers = []
                        
                        # Schedule for each section
                        for section in sections:
                            # Find available subject for this section
                            assigned = False
                            
                            # Shuffle subjects for variety
                            shuffled_subjects = standard_subjects.copy()
                            random.shuffle(shuffled_subjects)
                            
                            for subject in shuffled_subjects:
                                # Skip if we want to avoid same subject in parallel sections
                                if self.avoid_same_subject_var.get() and subject in assigned_subjects:
                                    continue
                                
                                available_teachers = self.subject_teachers.get(subject, [])
                                if not available_teachers:
                                    continue
                                
                                # Advanced balanced workload with teacher-specific daily hour constraints
                                teachers_by_workload = []
                                for teacher in available_teachers:
                                    if teacher not in assigned_teachers and time_slot not in teacher_schedule[teacher][day]:
                                        # Check hard constraints
                                        teacher_daily_periods = len([p for p in teacher_schedule[teacher][day]])
                                        teacher_weekly_periods = sum(len(teacher_schedule[teacher][d]) for d in self.days)
                                        
                                        # Convert periods to hours (35 minutes per period)
                                        teacher_daily_hours = teacher_daily_periods * (35 / 60)
                                        period_duration_hours = 35 / 60  # One period = 0.583 hours
                                        
                                        # Get teacher's min/max daily hours if set
                                        teacher_config = self.teacher_hours.get(teacher, {})
                                        min_daily_hours = teacher_config.get('min_hours', 3)  # Default 3 hours
                                        max_daily_hours = teacher_config.get('max_hours', 5)  # Default 5 hours
                                        
                                        # Hard constraint: Do not exceed teacher's max daily hours AFTER adding this period
                                        if teacher_daily_hours + period_duration_hours > max_daily_hours:
                                            continue
                                        
                                        # Priority for teachers below minimum hours
                                        if teacher_daily_hours < min_daily_hours:
                                            # Boost priority for teachers who need more hours to reach minimum
                                            priority_boost = (min_daily_hours - teacher_daily_hours) * 100  # Strong boost
                                            workload_score = teacher_daily_periods * 10 + teacher_weekly_periods - priority_boost
                                        else:
                                            # Normal scoring for teachers at or above minimum
                                            workload_score = teacher_daily_periods * 10 + teacher_weekly_periods
                                        
                                        teachers_by_workload.append((teacher, workload_score, teacher_daily_periods, teacher_weekly_periods))
                                
                                # Sort by workload score (fewest periods first)
                                teachers_by_workload.sort(key=lambda x: x[1])
                                
                                # Select best teacher
                                for teacher, workload_score, daily_periods, weekly_periods in teachers_by_workload:
                                    day_schedules[section].append({
                                        'period': f'P{period_num}',
                                        'subject': subject,
                                        'teacher': teacher,
                                        'time': time_slot,
                                        'duration': slot_duration
                                    })
                                    
                                    # Update trackers
                                    teacher_schedule[teacher][day].append(time_slot)
                                    assigned_subjects.append(subject)
                                    assigned_teachers.append(teacher)
                                    periods_scheduled[section] += 1
                                    assigned = True
                                    break
                                
                                if assigned:
                                    break
                            
                            if not assigned:
                                # Assign any available subject-teacher pair with teacher-specific constraints
                                all_pairs = []
                                for subject in standard_subjects:
                                    available_teachers = self.subject_teachers.get(subject, [])
                                    if not available_teachers:
                                        continue
                                    
                                    for teacher in available_teachers:
                                        if teacher not in assigned_teachers and time_slot not in teacher_schedule[teacher][day]:
                                            # Check hard constraints
                                            teacher_daily_periods = len([p for p in teacher_schedule[teacher][day]])
                                            teacher_weekly_periods = sum(len(teacher_schedule[teacher][d]) for d in self.days)
                                            
                                            # Convert periods to hours (35 minutes per period)
                                            teacher_daily_hours = teacher_daily_periods * (35 / 60)
                                            period_duration_hours = 35 / 60  # One period = 0.583 hours
                                            
                                            # Get teacher's min/max daily hours if set
                                            teacher_config = self.teacher_hours.get(teacher, {})
                                            min_daily_hours = teacher_config.get('min_hours', 3)  # Default 3 hours
                                            max_daily_hours = teacher_config.get('max_hours', 5)  # Default 5 hours
                                            
                                            # Hard constraint: Do not exceed teacher's max daily hours AFTER adding this period
                                            if teacher_daily_hours + period_duration_hours > max_daily_hours:
                                                continue
                                            
                                            # Priority for teachers below minimum hours
                                            if teacher_daily_hours < min_daily_hours:
                                                # Boost priority for teachers who need more hours to reach minimum
                                                priority_boost = (min_daily_hours - teacher_daily_hours) * 100  # Strong boost
                                                workload_score = teacher_daily_periods * 10 + teacher_weekly_periods - priority_boost
                                            else:
                                                # Normal scoring for teachers at or above minimum
                                                workload_score = teacher_daily_periods * 10 + teacher_weekly_periods
                                            
                                            all_pairs.append((teacher, workload_score, teacher_daily_periods, teacher_weekly_periods, subject))
                                
                                # Sort by workload score (fewest periods first)
                                all_pairs.sort(key=lambda x: x[1])
                                
                                # Select best teacher
                                for teacher, workload_score, daily_periods, weekly_periods, subject in all_pairs:
                                    day_schedules[section].append({
                                        'period': f'P{period_num}',
                                        'subject': subject,
                                        'teacher': teacher,
                                        'time': time_slot,
                                        'duration': slot_duration
                                    })
                                    
                                    # Update trackers
                                    teacher_schedule[teacher][day].append(time_slot)
                                    assigned_subjects.append(subject)
                                    assigned_teachers.append(teacher)
                                    periods_scheduled[section] += 1
                                    assigned = True
                                    break
                        
                        # Only increment period number for teaching periods (not recess)
                        if not time_slot_info.get('is_recess', False):
                            period_num += 1
                    
                    # Add day schedules to timetable
                    for section in sections:
                        if section not in self.timetable_data:
                            self.timetable_data[section] = {}
                        self.timetable_data[section][day] = day_schedules[section]
                
                self.log_generation(f"  ✓ Completed {standard}")
            
            # Balance teacher workload if requested
            if self.balance_load_var.get():
                self.balance_teacher_workload_multiple(teacher_schedule)
            
            # Verify no conflicts
            self.verify_no_conflicts()
            
            # Display generation summary
            self.display_generation_summary_multiple(teacher_schedule)
            
            self.log_generation("\n" + "=" * 60)
            self.log_generation("✅ TIMETABLE GENERATION COMPLETE!")
            self.log_generation(f"Generated for {len(self.sections)} sections")
            
            # Update view combobox
            self.view_class_combo['values'] = self.sections
            if self.sections:
                self.view_class_var.set(self.sections[0])
            
            self.update_status("Timetable generated successfully")
            messagebox.showinfo("Success", f"Timetable generated for {len(self.sections)} sections!")
            
        except Exception as e:
            self.log_generation(f"\n✗ ERROR: {str(e)}")
            self.log_generation(traceback.format_exc())
            messagebox.showerror("Generation Error", f"Error generating timetable:\n{str(e)}")
    
    def balance_teacher_workload_multiple(self, teacher_schedule):
        """Balance teacher workload when multiple teachers can teach same subject"""
        self.log_generation("\nBalancing teacher workload...")
        
        # Calculate current workload
        teacher_load = {}
        for teacher in self.teachers:
            total_periods = 0
            for day in self.days:
                if teacher in teacher_schedule:
                    total_periods += len(teacher_schedule[teacher][day])
            teacher_load[teacher] = total_periods
        
        # Calculate average
        total_periods = sum(teacher_load.values())
        avg_load = total_periods // max(1, len(self.teachers))
        
        self.log_generation(f"Average workload: {avg_load} periods per teacher")
        
        # Try to balance by swapping where possible
        swap_attempts = 0
        max_attempts = 100
        
        while swap_attempts < max_attempts:
            swap_made = False
            
            # Find most overloaded and underloaded teachers
            sorted_teachers = sorted(teacher_load.items(), key=lambda x: x[1])
            underloaded = sorted_teachers[0][0]
            overloaded = sorted_teachers[-1][0]
            
            # Check if balancing is needed
            if teacher_load[overloaded] - teacher_load[underloaded] <= 1:
                break
            
            # Try to find a swap opportunity
            for day in self.days:
                for section in self.sections:
                    if section not in self.timetable_data or day not in self.timetable_data[section]:
                        continue
                    
                    for i, period in enumerate(self.timetable_data[section][day]):
                        if period['teacher'] == overloaded and period['subject'] not in ['Free Period', 'Recess']:
                            # Check if underloaded teacher can teach this subject
                            subject = period['subject']
                            available_teachers = self.subject_teachers.get(subject, [])
                            
                            if underloaded in available_teachers:
                                # Check if underloaded teacher is free at this time
                                time_slot = period['time']
                                if time_slot not in teacher_schedule[underloaded][day]:
                                    # Make the swap
                                    self.timetable_data[section][day][i]['teacher'] = underloaded
                                    
                                    # Update schedules
                                    teacher_schedule[overloaded][day].remove(time_slot)
                                    teacher_schedule[underloaded][day].append(time_slot)
                                    
                                    # Update loads
                                    teacher_load[overloaded] -= 1
                                    teacher_load[underloaded] += 1
                                    
                                    self.log_generation(f"  Swapped {subject} from {overloaded} to {underloaded} in {section}")
                                    swap_made = True
                                    break
                    
                    if swap_made:
                        break
                if swap_made:
                    break
            
            swap_attempts += 1
            if not swap_made:
                break
        
        self.log_generation(f"Completed balancing after {swap_attempts} attempts")
    
    def verify_no_conflicts(self):
        """Verify there are no teacher conflicts in the timetable"""
        self.log_generation("\nVerifying no conflicts...")
        
        conflicts = []
        teacher_slot_tracker = {}
        
        for section in self.timetable_data:
            for day in self.timetable_data[section]:
                for period in self.timetable_data[section][day]:
                    if period['teacher'] and period['subject'] not in ['Free Period', 'Recess']:
                        key = (period['teacher'], day, period['time'])
                        
                        if key in teacher_slot_tracker:
                            conflicts.append({
                                'teacher': period['teacher'],
                                'day': day,
                                'time': period['time'],
                                'section1': teacher_slot_tracker[key],
                                'section2': section,
                                'subject': period['subject']
                            })
                        else:
                            teacher_slot_tracker[key] = section
        
        if conflicts:
            self.log_generation("⚠ CONFLICTS FOUND:")
            for conflict in conflicts:
                self.log_generation(f"  {conflict['teacher']} scheduled at same time ({conflict['day']} {conflict['time']}) in {conflict['section1']} and {conflict['section2']}")
            return False
        else:
            self.log_generation("✓ No teacher conflicts found")
            return True
    
    def display_generation_summary_multiple(self, teacher_schedule):
        """Display generation summary for multiple teachers per subject"""
        self.log_generation("\n--- GENERATION SUMMARY ---")
        
        # Teacher workload
        self.log_generation("Teacher Workload:")
        for teacher in self.teachers:
            total_periods = 0
            for day in self.days:
                if teacher in teacher_schedule:
                    total_periods += len(teacher_schedule[teacher][day])
            
            # Count subjects this teacher teaches
            teacher_subjects = []
            for subject, teachers in self.subject_teachers.items():
                if teacher in teachers:
                    teacher_subjects.append(subject)
            
            self.log_generation(f"  {teacher}: {total_periods} periods ({len(teacher_subjects)} subjects)")
        
        # Teacher daily breakdown
        self.log_generation("\nTeacher Daily Teaching Hours:")
        for teacher in self.teachers:
            self.log_generation(f"\n{teacher}:")
            daily_total = 0
            for day in self.days:
                day_periods = len(teacher_schedule[teacher][day]) if teacher in teacher_schedule else 0
                daily_total += day_periods
                hours = day_periods * self.config.get('period_duration', 35) / 60  # Convert to hours
                self.log_generation(f"  {day}: {day_periods} periods ({hours:.1f} hours)")
            self.log_generation(f"  Total: {daily_total} periods ({daily_total * self.config.get('period_duration', 35) / 60:.1f} hours)")
        
        # Subject distribution
        self.log_generation("\nSubject Distribution by Teacher:")
        for subject, teachers in self.subject_teachers.items():
            self.log_generation(f"  {subject}: {len(teachers)} teachers ({', '.join(teachers)})")
    
    def regenerate_all(self):
        """Regenerate all timetables"""
        if messagebox.askyesno("Confirm", "Regenerate all timetables? This will overwrite existing data."):
            self.timetable_data.clear()
            self.gen_status_text.delete(1.0, tk.END)
            self.generate_timetable()
    
    # ========== VIEW METHODS ==========
    
    def view_text_timetable(self):
        """Display timetable in text format"""
        section = self.view_class_var.get()
        day = self.view_day_var.get()
        
        if not section:
            messagebox.showerror("Error", "Please select a class/section")
            return
        
        if section not in self.timetable_data:
            messagebox.showerror("Error", f"No timetable generated for {section}")
            return
        
        self.view_text.delete(1.0, tk.END)
        
        # Display header
        header = f"{'='*70}\n"
        header += f"TIMETABLE FOR {section}\n"
        header += f"{'='*70}\n\n"
        self.view_text.insert(tk.END, header)
        
        if day == "All":
            # Display all days
            for day_name in self.days:
                if day_name in self.timetable_data[section]:
                    self.display_day_timetable(section, day_name)
        else:
            # Display specific day
            if day in self.timetable_data[section]:
                self.display_day_timetable(section, day)
            else:
                self.view_text.insert(tk.END, f"No timetable for {day}\n")
    
    def display_day_timetable(self, section, day):
        """Display timetable for a specific day"""
        day_data = self.timetable_data[section][day]
        
        # Debug: Log what periods are in data
        self.log_generation(f"DEBUG: Displaying {len(day_data)} periods for {section} {day}")
        for i, period in enumerate(day_data):
            self.log_generation(f"DEBUG: Period {i+1}: {period['time']} - {period['period']} - {period['subject']} ({period['duration']}min)")
        
        self.view_text.insert(tk.END, f"\n{'='*70}\n")
        self.view_text.insert(tk.END, f"{day.upper()}\n")
        self.view_text.insert(tk.END, f"{'='*70}\n")
        
        if not day_data:
            self.view_text.insert(tk.END, "No classes scheduled\n")
            return
        
        # Create formatted table with end times
        table_header = f"{'Time':<10} {'End Time':<10} {'Period':<10} {'Subject':<25} {'Teacher':<20}\n"
        table_header += "-" * 80 + "\n"
        self.view_text.insert(tk.END, table_header)
        
        # Debug: Count displayed periods
        displayed_count = 0
        
        for period in day_data:
            displayed_count += 1
            self.log_generation(f"DEBUG: Displaying period {displayed_count}: {period['time']} - {period['period']} - {period['subject']}")
            
            # Calculate end time
            start_time = datetime.strptime(period['time'], "%H:%M")
            end_time = start_time + timedelta(minutes=period['duration'])
            end_time_str = end_time.strftime("%H:%M")
            
            if period['period'] == 'R' or period['subject'] == 'Recess':
                line = f"{period['time']:<10} {end_time_str:<10} {'RECESS':<10} {period['subject']:<25} {'':<20}\n"
            else:
                line = f"{period['time']:<10} {end_time_str:<10} {period['period']:<10} {period['subject'][:23]:<25} {period['teacher'][:18]:<20}\n"
            self.view_text.insert(tk.END, line)
        
        self.log_generation(f"DEBUG: Total periods displayed: {displayed_count}")
    
    def view_teacher_schedule(self):
        """Display teacher schedule"""
        self.teacher_text.delete(1.0, tk.END)
        
        if not self.timetable_data:
            self.teacher_text.insert(tk.END, "No timetable generated yet.")
            return
        
        # Collect teacher schedules
        teacher_schedules = {}
        for section in self.timetable_data:
            for day in self.timetable_data[section]:
                for period in self.timetable_data[section][day]:
                    if period['teacher'] and period['subject'] not in ['Free Period', 'Recess']:
                        teacher = period['teacher']
                        if teacher not in teacher_schedules:
                            teacher_schedules[teacher] = {}
                        if day not in teacher_schedules[teacher]:
                            teacher_schedules[teacher][day] = []
                        teacher_schedules[teacher][day].append({
                            'time': period['time'],
                            'subject': period['subject'],
                            'section': section,
                            'period': period['period']
                        })
        
        # Display schedules
        for teacher, days in teacher_schedules.items():
            self.teacher_text.insert(tk.END, f"\n{'='*70}\n")
            self.teacher_text.insert(tk.END, f"SCHEDULE FOR {teacher}\n")
            self.teacher_text.insert(tk.END, f"{'='*70}\n")
            
            for day in self.days:
                if day in days:
                    self.teacher_text.insert(tk.END, f"\n{day.upper()}:\n")
                    self.teacher_text.insert(tk.END, f"{'Time':<10} {'Period':<10} {'Subject':<20} {'Section':<15}\n")
                    self.teacher_text.insert(tk.END, "-" * 60 + "\n")
                    
                    for period in sorted(days[day], key=lambda x: x['time']):
                        self.teacher_text.insert(tk.END, f"{period['time']:<10} {period['period']:<10} {period['subject'][:18]:<20} {period['section']:<15}\n")
    
    # ========== EXPORT METHODS ==========
    
    def create_master_timetable_sheet(self, writer):
        """Create a master timetable sheet with all sections using professional formatting"""
        try:
            # Define professional styles
            header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            center_alignment = Alignment(horizontal="center", vertical="center")
            border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            title_font = Font(size=16, bold=True, color="FFFFFF")
            subtitle_font = Font(size=11, bold=True, color="366092")
            
            # Get all unique periods and their times
            all_periods = []
            for section in self.sections:
                if section in self.timetable_data:
                    for day in self.days:
                        if day in self.timetable_data[section]:
                            for period in self.timetable_data[section][day]:
                                if period['period'] not in [p['period'] for p in all_periods]:
                                    all_periods.append({
                                        'period': period['period'],
                                        'time': period['time'],
                                        'duration': period['duration']
                                    })
            
            # Sort periods by time
            all_periods.sort(key=lambda x: datetime.strptime(x['time'], "%H:%M"))
            
            # Group sections by standard
            sections_by_standard = {}
            for section in self.sections:
                standard = section.split('-')[0]
                if standard not in sections_by_standard:
                    sections_by_standard[standard] = []
                sections_by_standard[standard].append(section)
            
            # Create master timetable for each standard
            for standard, sections in sections_by_standard.items():
                ws = writer.book.create_sheet(title=f"{standard}_Master_Timetable"[:31])
                
                # Add title
                title_range = f"A1:{chr(ord('A') + len(sections) + 2)}1"
                ws.merge_cells(title_range)
                ws['A1'] = f"MASTER TIMETABLE - {standard}"
                ws['A1'].font = title_font
                ws['A1'].fill = header_fill
                ws['A1'].alignment = center_alignment
                
                # Add school timing subtitle
                subtitle_range = f"A2:{chr(ord('A') + len(sections) + 2)}2"
                ws.merge_cells(subtitle_range)
                ws['A2'] = f"Timing: {self.school_timing.get('start', 'N/A')} to {self.school_timing.get('end', 'N/A')}"
                ws['A2'].font = subtitle_font
                ws['A2'].alignment = Alignment(horizontal="center")
                
                # Create headers
                headers = ['Period', 'Time', 'Duration'] + sections
                for col_idx, header in enumerate(headers, start=1):
                    cell = ws.cell(row=4, column=col_idx, value=header)
                    cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
                    cell.font = header_font
                    cell.alignment = center_alignment
                    cell.border = border
                
                # Define day colors for variety
                day_colors = {
                    'Monday': 'E2EFDA',
                    'Tuesday': 'FFF2CC', 
                    'Wednesday': 'FCE4D6',
                    'Thursday': 'EDEDED',
                    'Friday': 'D9EAD3',
                    'Saturday': 'F4CCCC'
                }
                
                # Fill timetable data
                row = 5
                for period_info in all_periods:
                    if period_info['period'] == 'R':
                        continue  # Skip recess in master grid
                    
                    # Period info
                    ws.cell(row=row, column=1, value=period_info['period'])
                    ws.cell(row=row, column=2, value=period_info['time'])
                    ws.cell(row=row, column=3, value=f"{period_info['duration']}min")
                    
                    # Apply borders to period info
                    for col in range(1, 4):
                        ws.cell(row=row, column=col).border = border
                    
                    # Section data
                    for col_idx, section in enumerate(sections, start=4):
                        cell_content = ''
                        cell_fill = None
                        
                        if section in self.timetable_data:
                            for day in self.days:
                                if day in self.timetable_data[section]:
                                    for period in self.timetable_data[section][day]:
                                        if period['period'] == period_info['period']:
                                            if period['subject'] == 'Recess':
                                                cell_content = 'RECESS'
                                                cell_fill = PatternFill(start_color="C4D79B", end_color="C4D79B", fill_type="solid")
                                            else:
                                                cell_content = f"{period['subject']}\n{period['teacher']}"
                                                # Apply day color to first column of each section
                                                if col_idx == 4:  # First section column
                                                    day_color = day_colors.get(day, 'FFFFFF')
                                                    cell_fill = PatternFill(start_color=day_color, end_color=day_color, fill_type="solid")
                                            break
                        
                        cell = ws.cell(row=row, column=col_idx, value=cell_content)
                        cell.border = border
                        if cell_fill:
                            cell.fill = cell_fill
                        cell.alignment = center_alignment
                    
                    row += 1
                
                # Auto-size columns
                for col_idx in range(1, len(headers) + 1):
                    col_letter = chr(ord('A') + col_idx - 1)
                    ws.column_dimensions[col_letter].width = 15
                
                self.log_export(f"✓ Professional master timetable created for {standard}")
            
        except Exception as e:
            self.log_export(f"✗ Error creating master timetable: {str(e)}")
    
    def export_excel(self):
        """Export all timetables to Excel"""
        if not self.timetable_data:
            messagebox.showerror("Error", "No timetable data to export. Please generate timetable first.")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile="timetable_output.xlsx"
        )
        
        if filename:
            try:
                self.log_export(f"Exporting to {filename}...")
                
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    # Save configuration
                    config_data = []
                    if self.config:
                        config_data.append(['School Timing', f"{self.school_timing.get('start', 'N/A')} to {self.school_timing.get('end', 'N/A')}"])
                        config_data.append(['Period Duration', f"{self.config.get('period_duration', 'N/A')} minutes"])
                        config_data.append(['Period Overtime', f"{self.config.get('period_overtime', 'N/A')} minutes"])
                        config_data.append(['Teaching Hours/Day', f"{self.config.get('teaching_hours', 'N/A')}"])
                        config_data.append(['Working Days', ', '.join(self.days)])
                        
                        for i, recess in enumerate(self.config.get('recess_times', [])):
                            config_data.append([f'Recess {i+1}', f"{recess['name']} at {recess['start']} ({recess['duration']}min)"])
                    
                    if config_data:
                        config_df = pd.DataFrame(config_data, columns=['Parameter', 'Value'])
                        config_df.to_excel(writer, sheet_name='Configuration', index=False)
                        self.log_export("✓ Configuration sheet created")
                    
                    # Create professional timetable format for each section
                    sections_exported = 0
                    for section in self.sections:
                        if section in self.timetable_data:
                            # Define professional styles
                            header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                            header_font = Font(color="FFFFFF", bold=True)
                            center_alignment = Alignment(horizontal="center", vertical="center")
                            border = Border(
                                left=Side(style='thin'),
                                right=Side(style='thin'),
                                top=Side(style='thin'),
                                bottom=Side(style='thin')
                            )
                            title_font = Font(size=16, bold=True, color="FFFFFF")
                            subtitle_font = Font(size=11, bold=True, color="366092")
                            
                            # Create worksheet
                            ws = writer.book.create_sheet(title=f"{section}_Timetable"[:31])
                            
                            # Add title
                            title_range = f"A1:{chr(ord('A') + len(self.days) + 2)}1"
                            ws.merge_cells(title_range)
                            ws['A1'] = f"TIMETABLE - {section}"
                            ws['A1'].font = title_font
                            ws['A1'].fill = header_fill
                            ws['A1'].alignment = center_alignment
                            
                            # Add school timing subtitle
                            subtitle_range = f"A2:{chr(ord('A') + len(self.days) + 2)}2"
                            ws.merge_cells(subtitle_range)
                            ws['A2'] = f"Timing: {self.school_timing.get('start', 'N/A')} to {self.school_timing.get('end', 'N/A')}"
                            ws['A2'].font = subtitle_font
                            ws['A2'].alignment = Alignment(horizontal="center")
                            
                            # Get all unique periods and their times
                            all_periods = []
                            for day in self.days:
                                if day in self.timetable_data[section]:
                                    for period in self.timetable_data[section][day]:
                                        if period['period'] not in [p['period'] for p in all_periods]:
                                            all_periods.append({
                                                'period': period['period'],
                                                'time': period['time'],
                                                'duration': period['duration']
                                            })
                            
                            # Sort periods by time
                            all_periods.sort(key=lambda x: datetime.strptime(x['time'], "%H:%M"))
                            
                            # Create headers
                            headers = ['Period', 'Time', 'Duration'] + self.days
                            for col_idx, header in enumerate(headers, start=1):
                                cell = ws.cell(row=4, column=col_idx, value=header)
                                cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
                                cell.font = header_font
                                cell.alignment = center_alignment
                                cell.border = border
                            
                            # Define day colors for variety
                            day_colors = {
                                'Monday': 'E2EFDA',
                                'Tuesday': 'FFF2CC', 
                                'Wednesday': 'FCE4D6',
                                'Thursday': 'EDEDED',
                                'Friday': 'D9EAD3',
                                'Saturday': 'F4CCCC'
                            }
                            
                            # Fill timetable data
                            row = 5
                            for period_info in all_periods:
                                if period_info['period'] == 'R':
                                    continue  # Skip recess in grid
                                
                                # Period info
                                ws.cell(row=row, column=1, value=period_info['period'])
                                ws.cell(row=row, column=2, value=period_info['time'])
                                ws.cell(row=row, column=3, value=f"{period_info['duration']}min")
                                
                                # Apply borders to period info
                                for col in range(1, 4):
                                    ws.cell(row=row, column=col).border = border
                                
                                # Day-wise data
                                for col_idx, day in enumerate(self.days, start=4):
                                    cell_content = ''
                                    cell_fill = None
                                    
                                    if day in self.timetable_data[section]:
                                        for period in self.timetable_data[section][day]:
                                            if period['period'] == period_info['period']:
                                                if period['subject'] == 'Recess':
                                                    cell_content = 'RECESS'
                                                    cell_fill = PatternFill(start_color="C4D79B", end_color="C4D79B", fill_type="solid")
                                                else:
                                                    cell_content = f"{period['subject']}\n{period['teacher']}"
                                                    # Apply day color
                                                    day_color = day_colors.get(day, 'FFFFFF')
                                                    cell_fill = PatternFill(start_color=day_color, end_color=day_color, fill_type="solid")
                                                break
                                    
                                    cell = ws.cell(row=row, column=col_idx, value=cell_content)
                                    cell.border = border
                                    if cell_fill:
                                        cell.fill = cell_fill
                                    cell.alignment = center_alignment
                                
                                row += 1
                            
                            # Auto-size columns
                            for col_idx in range(1, len(headers) + 1):
                                col_letter = chr(ord('A') + col_idx - 1)
                                ws.column_dimensions[col_letter].width = 15
                            
                            sections_exported += 1
                            self.log_export(f"✓ Professional timetable sheet created for {section}")
                            self.log_export(f"✓ Timetable sheet created for {section}")
                    
                    # Create summary timetable with all sections in one sheet
                    self.create_master_timetable_sheet(writer)
                    
                    # Save master data
                    master_data = {
                        'Subjects': self.subjects,
                        'Standards': self.standards,
                        'Sections': self.sections,
                        'Teachers': self.teachers
                    }
                    
                    master_df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in master_data.items()]))
                    master_df.to_excel(writer, sheet_name='Master_Data', index=False)
                    
                    # Save subject-teacher assignments (multiple teachers per subject)
                    assignments_data = []
                    for subject, teachers in self.subject_teachers.items():
                        for teacher in teachers:
                            assignments_data.append({'Subject': subject, 'Teacher': teacher})
                    
                    for standard, subjects in self.standard_subjects.items():
                        for subject in subjects:
                            assignments_data.append({'Standard': standard, 'Subject': subject})
                    
                    if assignments_data:
                        assignments_df = pd.DataFrame(assignments_data)
                        assignments_df.to_excel(writer, sheet_name='Assignments', index=False)
                    
                    self.log_export("✓ Master data and assignments sheets created")
                
                self.log_export(f"\n✅ EXPORT COMPLETE!")
                self.log_export(f"Exported {sections_exported} sections to Excel")
                self.update_status(f"Exported to {filename}")
                messagebox.showinfo("Success", f"Timetable exported to {filename}")
                
            except Exception as e:
                self.log_export(f"\n✗ ERROR: {str(e)}")
                messagebox.showerror("Export Error", f"Error exporting to Excel:\n{str(e)}")
    
    def export_excel_selected(self):
        """Export selected sections to Excel"""
        # For now, just export all
        self.export_excel()
    
    def export_summary(self):
        """Export summary report"""
        if not self.timetable_data:
            messagebox.showerror("Error", "No timetable data to export. Please generate timetable first.")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialfile="timetable_summary.txt"
        )
        
        if filename:
            try:
                with open(filename, 'w') as f:
                    f.write("=" * 70 + "\n")
                    f.write("TIMETABLE GENERATION SUMMARY\n")
                    f.write("=" * 70 + "\n\n")
                    
                    f.write(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                    
                    f.write("CONFIGURATION:\n")
                    f.write("-" * 40 + "\n")
                    f.write(f"School Timing: {self.school_timing.get('start', 'N/A')} to {self.school_timing.get('end', 'N/A')}\n")
                    f.write(f"Working Days: {', '.join(self.days)}\n")
                    f.write(f"Period Duration: {self.config.get('period_duration', 'N/A')} minutes\n")
                    f.write(f"Teaching Hours/Day: {self.config.get('teaching_hours', 'N/A')}\n\n")
                    
                    f.write("STATISTICS:\n")
                    f.write("-" * 40 + "\n")
                    f.write(f"Subjects: {len(self.subjects)}\n")
                    f.write(f"Teachers: {len(self.teachers)}\n")
                    f.write(f"Standards: {len(self.standards)}\n")
                    f.write(f"Sections: {len(self.sections)}\n\n")
                    
                    f.write("SUBJECT-TEACHER ASSIGNMENTS:\n")
                    f.write("-" * 40 + "\n")
                    for subject, teachers in self.subject_teachers.items():
                        f.write(f"{subject}: {len(teachers)} teachers ({', '.join(teachers)})\n")
                    
                    f.write("\nTEACHER WORKLOAD:\n")
                    f.write("-" * 40 + "\n")
                    
                    # Calculate teacher workload
                    teacher_periods = {}
                    for section in self.timetable_data:
                        for day in self.timetable_data[section]:
                            for period in self.timetable_data[section][day]:
                                if period['teacher'] and period['subject'] not in ['Free Period', 'Recess']:
                                    teacher = period['teacher']
                                    teacher_periods[teacher] = teacher_periods.get(teacher, 0) + 1
                    
                    for teacher in sorted(self.teachers):
                        periods = teacher_periods.get(teacher, 0)
                        subjects = [s for s, teachers in self.subject_teachers.items() if teacher in teachers]
                        f.write(f"{teacher}: {periods} periods ({len(subjects)} subjects)\n")
                
                self.log_export(f"✓ Summary report exported to {filename}")
                self.update_status(f"Summary exported")
                messagebox.showinfo("Success", f"Summary report exported to {filename}")
                
            except Exception as e:
                self.log_export(f"✗ ERROR: {str(e)}")
                messagebox.showerror("Export Error", f"Error exporting summary:\n{str(e)}")
    
    def save_configuration(self):
        """Save all configuration to JSON file"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialfile="timetable_config.json"
        )
        
        if filename:
            try:
                data = {
                    'config': self.config,
                    'subjects': self.subjects,
                    'standards': self.standards,
                    'sections': self.sections,
                    'teachers': self.teachers,
                    'subject_teachers': self.subject_teachers,  # Now subject: [teachers]
                    'standard_subjects': self.standard_subjects,
                    'timetable_data': self.timetable_data,
                    'school_timing': self.school_timing,
                    'days': self.days
                }
                
                with open(filename, 'w') as f:
                    json.dump(data, f, indent=4, default=str)
                
                self.log_export(f"✓ Configuration saved to {filename}")
                self.update_status(f"Configuration saved")
                messagebox.showinfo("Success", f"Configuration saved to {filename}")
                
            except Exception as e:
                self.log_export(f"✗ ERROR: {str(e)}")
                messagebox.showerror("Save Error", f"Error saving configuration:\n{str(e)}")
    
    def load_configuration(self):
        """Load configuration from JSON file"""
        filename = filedialog.askopenfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'r') as f:
                    data = json.load(f)
                
                # Load data
                self.config = data.get('config', {})
                self.subjects = data.get('subjects', [])
                self.standards = data.get('standards', [])
                self.sections = data.get('sections', [])
                self.teachers = data.get('teachers', [])
                
                # Handle both old (subject: teacher) and new (subject: [teachers]) format
                subject_teachers_data = data.get('subject_teachers', {})
                self.subject_teachers = {}
                for subject, teacher_data in subject_teachers_data.items():
                    if isinstance(teacher_data, list):
                        self.subject_teachers[subject] = teacher_data
                    else:
                        # Convert old format to new format
                        self.subject_teachers[subject] = [teacher_data]
                
                self.standard_subjects = data.get('standard_subjects', {})
                self.timetable_data = data.get('timetable_data', {})
                self.school_timing = data.get('school_timing', {})
                self.days = data.get('days', self.days)
                
                # Update GUI widgets
                self.update_gui_from_data()
                
                self.log_export(f"✓ Configuration loaded from {filename}")
                self.update_status(f"Configuration loaded")
                messagebox.showinfo("Success", f"Configuration loaded from {filename}")
                
            except Exception as e:
                self.log_export(f"✗ ERROR: {str(e)}")
                messagebox.showerror("Load Error", f"Error loading configuration:\n{str(e)}")
    
    def update_gui_from_data(self):
        """Update GUI widgets from loaded data"""
        # Update subjects listbox
        self.subjects_listbox.delete(0, tk.END)
        for subject in self.subjects:
            self.subjects_listbox.insert(tk.END, subject)
        
        # Update standards listbox
        self.std_listbox.delete(0, tk.END)
        for standard in self.standards:
            self.std_listbox.insert(tk.END, standard)
        
        # Update sections tree
        for item in self.sec_tree.get_children():
            self.sec_tree.delete(item)
        for section in self.sections:
            if '-' in section:
                std, sec = section.split('-', 1)
                self.sec_tree.insert('', 'end', values=(std, sec))
        
        # Update teachers listbox
        self.teachers_listbox.delete(0, tk.END)
        for teacher in self.teachers:
            self.teachers_listbox.insert(tk.END, teacher)
        
        # Update recess tree
        for item in self.recess_tree.get_children():
            self.recess_tree.delete(item)
        if 'recess_times' in self.config:
            for recess in self.config['recess_times']:
                self.recess_tree.insert('', 'end', values=(recess['name'], recess['start'], f"{recess['duration']} min"))
        
        # Update comboboxes
        self.std_combo['values'] = self.standards
        self.ts_teacher_combo['values'] = self.teachers
        self.cs_class_combo['values'] = self.standards
        self.view_class_combo['values'] = self.sections
        
        # Update assignment trees
        self.update_assignment_trees()
        
        # Update day checkboxes
        for day, var in self.day_vars.items():
            var.set(day in self.days)
        
        # Refresh assignment displays
        if self.ts_teacher_var.get():
            self.on_teacher_selected()
        if self.cs_class_var.get():
            self.on_class_selected()
    
    def show_help(self):
        """Show help dialog"""
        help_text = """
        Advanced Timetable Generator - Help Guide
        
        HOW TO CREATE A TIMETABLE:
        
        1. School Timing (First Tab):
           - Set school start and end times (HH:MM format)
           - Select working days (Monday to Saturday)
           - Click "Save School Timing"
        
        2. Periods (Second Tab):
           - Set period duration (e.g., 45 minutes)
           - Set period overtime (e.g., 5 minutes between periods)
           - Set teaching hours per day (e.g., 7)
           - Add recess periods with name, start time, and duration
           - Click "Save Period Settings"
        
        3. Subjects (Third Tab):
           - Add all subjects to be taught
           - Subjects can be removed or cleared
        
        4. Classes (Fourth Tab):
           - Add standards/classes (e.g., "1", "2", "3")
           - For each standard, add sections (e.g., "A", "B", "C")
           - Sections will be created as "1-A", "1-B", etc.
        
        5. Teachers (Fifth Tab):
           - Add all teachers
        
        6. Assignments (Sixth Tab):
           - Teacher-Subject Tab:
             * Select a teacher
             * Check subjects they teach (multiple allowed)
             * Click "Save Teacher Subjects"
             * MULTIPLE TEACHERS CAN TEACH SAME SUBJECT
           - Class-Subject Tab:
             * Select a class/standard
             * Check subjects to be taught in that class
             * Click "Save Class Subjects"
        
        7. Generate (Seventh Tab):
           - Click "Validate Data" to check everything
           - Configure generation options:
             * Avoid Teacher Conflicts: Prevents same teacher in two classes at same time
             * Balance Teacher Workload: Distributes periods evenly
             * Avoid Same Subject in Parallel Sections: Different subjects in same standard sections
           - Click "Generate Timetable"
        
        8. View (Eighth Tab):
           - Select a section and day to view timetable
           - View teacher schedules
        
        9. Export (Ninth Tab):
           - Export to Excel (all data)
           - Save/Load configuration as JSON
           - Export summary report
        
        KEY FEATURES:
        - MULTIPLE TEACHERS PER SUBJECT
        - Conflict-free scheduling
        - Balanced teacher workload
        - Smart subject distribution
        - Full data validation
        - Export to multiple formats
        
        EXAMPLE:
        - Teacher A can teach Maths and Science
        - Teacher B can teach Maths and Gujarati
        - Result: Maths has both Teacher A and Teacher B assigned
        - Timetable generator can use either teacher for Maths classes
        
        TIPS:
        - Fill tabs in order
        - Always validate before generating
        - Save configuration frequently
        - Check teacher schedules for conflicts
        
        For issues or suggestions, please report them.
        """
        
        help_window = tk.Toplevel(self.root)
        help_window.title("Help Guide")
        help_window.geometry("700x600")
        
        text_widget = scrolledtext.ScrolledText(help_window, wrap=tk.WORD)
        text_widget.insert(tk.END, help_text)
        text_widget.config(state='disabled')
        text_widget.pack(fill='both', expand=True, padx=10, pady=10)
        
        ttk.Button(help_window, text="Close", 
                  command=help_window.destroy).pack(pady=10)

def main():
    """Main function to run the GUI application"""
    root = tk.Tk()
    
    # Try to set theme
    try:
        import ttkthemes
        style = ttkthemes.ThemedStyle(root)
        style.set_theme("arc")
    except ImportError:
        pass  # Use default theme if ttkthemes not available
    
    # Create custom style
    style = ttk.Style()
    style.configure('Accent.TButton', font=('Arial', 11, 'bold'))
    
    app = TimetableGUI(root)
    root.mainloop()

if __name__ == "__main__":
    # Check required packages
    try:
        import pandas as pd
        from openpyxl import Workbook
    except ImportError as e:
        print(f"Error: {e}")
        print("\nPlease install required packages:")
        print("pip install pandas openpyxl")
        exit(1)
    
    main()