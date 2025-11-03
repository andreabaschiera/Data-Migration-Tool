def month_name_to_number(month_name):
    # Dictionary mapping month names to numbers
    months = {
        "January": 1, "February": 2, "March": 3, "April": 4,
        "May": 5, "June": 6, "July": 7, "August": 8,
        "September": 9, "October": 10, "November": 11, "December": 12
    }
    # Return the corresponding number or None if not found
    return months.get(month_name.capitalize(), None)