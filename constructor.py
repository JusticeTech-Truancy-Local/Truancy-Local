# Author: Lindsey Allen
# Date: March 3, 2025

# This is the constructor for the Student class, to be able to create objects with all the relevant attributes from the Progress Book PDF.
# It has a construct_empty() method so that a Student object can be created and even added to a data structure without having all the attributes
# at initialization.

class Student:

    # this is the maximum threshold for truancy hours. when reached, unexcused absences will be marked red in the spreadsheet
    redThreshold = 40

    def __init__(self, id, firstName, lastName, age, grade, excused=0, unexcused=0, medical=0, suspension=0, schoolTotal=0, attendingTotal=0, absenceTotal=0, state="NOT_TRUANT"):
        # primary key; student's school ID number; six digit number
        self.id = id
        self.firstName = firstName
        self.lastName = lastName
        self.age = age
        # grade level (ex: KG, 01, 05)
        self.grade = grade

        # Excused Absence Hours (does not include medically included)
        self.excused = excused 

        # Unexcused Absence Hours 
        self.unexcused = unexcused

        # Medically Excused Absenses
        self.medical = medical

        # Suspension Hours 
        self.suspension = suspension

        # Total School Hours (Hours of School Time in the Time Period)
        self.schoolTotal = schoolTotal

        # Attending Total is the number of hours, out of Total School Hours, the student has attended
        # Attending Hours = Total School Hours - (Excused Hours + Unexcused Hours + Medically Excused Hours + Suspension Hours)
        self.attendingTotal = attendingTotal

        # Total Absence Hours = Unexcused Absence Hours + Excused Absence Hours
        self.absenceTotal = absenceTotal
        
        # State = status of the student in the truancy process
        self.state = state
    
    @classmethod
    def create_empty(cls):
        return cls(None, None, None, None, None, None, None, None, None, None, None, None)
        
        
    #simple method to print all data in a student object
    def print(self):
    	print(str(self.id) + "\t" + self.firstName + " " + self.lastName + "\t\t" + str(self.age) + "\t" + str(self.grade) + "\t" + str(self.excused) + "\t" + str(self.unexcused) + "\t\t" + str(self.medical) + "\t" + str(self.suspension) + "\t\t" + str(self.schoolTotal) + "\t\t" + str(self.attendingTotal) + "\t\t" + str(self.absenceTotal))
    	
    #prints the headers to indicate what data corresponds to what field in the print method
    @staticmethod
    def printHeaders():
    	print("id\tfirstName lastName\tage\tgrade\texcused\tunexcused\tmedical\tsuspension\tschoolTotal\tattendingTotal\tabsenceTotal")
