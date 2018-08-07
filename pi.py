from datetime import timedelta, datetime
from win32com.client import Dispatch

def pi_read(server, tag, start=None, end=None, frequency=None):

    pisdk = Dispatch('PISDK.PISDK')
    my_server = pisdk.Servers(server)

    time_start = Dispatch('PITimeServer.PITimeFormat')
    time_end = Dispatch('PITimeServer.PITimeFormat')
    
    sample_point = my_server.PIPoints[tag]
    
    uom = sample_point.PointAttributes.Item("EngUnits").Value
    description = sample_point.PointAttributes.Item('Descriptor').Value 
    
    if start != None and end != None:
        time_start.InputString = start.strftime('%Y-%m-%d %H:%M:%S')
        time_end.InputString = end.strftime('%Y-%m-%d %H:%M:%S')
        sample_values = sample_point.Data.Summaries2(time_start, time_end, frequency, 5, 0, None)
        values = sample_values('Average').Value
        data = [x.Value for x in values]
    elif start != None and end == None:
        end = start + timedelta(seconds=1)
        time_start.InputString = start.strftime('%m-%d-%Y %H:%M:%S')
        time_end.InputString = end.strftime('%m-%d-%Y %H:%M:%S')
        sample_values = sample_point.Data.Summaries2(time_start, time_end, frequency, 5, 0, None)
        values = sample_values('Average').Value
        data = [x.Value for x in values][0]
    else:
        data = sample_point.data.Snapshot.Value  
        
    return tag, description, uom, data

def read_batch(server, taglist, start, end, interval):

    for tag in taglist:
        try: yield pi_read(server, tag, start, end, interval)
        except: pass
