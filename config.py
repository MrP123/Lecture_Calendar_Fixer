def get_travel_time(location: str) -> int:
    """Get travel time in minutes from MCI IV to the specified MCI location"""
    
    mci_travel_times = {
        "MCI I": 15,
        "MCI II": 15,
        "MCI III": 35,
        "MCI IV": 0,
        "MCI V": 20,
        "MCI VI": 5,
    }
    return mci_travel_times.get(location, 0)