
def is_async_online_lecture(subject:str, room: str) -> bool:
    """Check if an event is an asynchronous online lecutre"""
    return "Geleitetes Selbststudium" in subject and room == "Online"

def at_different_location(location: str) -> bool:
    """Check if the event is at a different location than MCI IV"""
    return location != "MCI IV"

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