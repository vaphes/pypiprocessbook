from datetime import timedelta, datetime
from win32com.client import Dispatch

class PI:

    def __init__(self, server):
        try:
            pisdk = Dispatch('PISDK.PISDK')
            self.server =  pisdk.Servers(server)
            self.server_name = server
        except:
            raise ValueError('Server {} not found or could not connect!'.format(server))

    def read(self, tag, start=None, end=None, frequency=60):

        #pisdk = Dispatch('PISDK.PISDK')
        #my_server = pisdk.Servers(self.server)

        time_start = Dispatch('PITimeServer.PITimeFormat')
        time_end = Dispatch('PITimeServer.PITimeFormat')

        if start:
            if isinstance(start, str):
                start = datetime.strptime(start, '%Y-%m-%d %H:%M:%S')
            elif not isinstance(start, datetime):
                raise ValueError('Start date and time must be str ou datetime object!')
        
        if end:
            if isinstance(end, str):
                end = datetime.strptime(end, '%Y-%m-%d %H:%M:%S')
            elif not isinstance(end, datetime):
                raise ValueError('End date and time must be str ou datetime object!')

        #sample_point = my_server.PIPoints[tag]
        try:
            sample_point = self.server.PIPoints[tag]
        except:
            raise ValueError('TAG {} does not exist on {}!'.format(tag, self.server_name))
        
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
            time_start.InputString = start.strftime('%Y-%m-%d %H:%M:%S')
            time_end.InputString = end.strftime('%Y-%m-%d %H:%M:%S')
            sample_values = sample_point.Data.Summaries2(time_start, time_end, frequency, 5, 0, None)
            values = sample_values('Average').Value
            data = [x.Value for x in values][0]
        else:
            data = sample_point.data.Snapshot.Value  
            
        return tag, description, data, uom

    def read_batch(self, taglist, start, end, interval):

        for tag in taglist:
            try: 
                yield pi_read(tag, start, end, interval)
            except: 
                print('tag {} could not be retrieved!'.format(tag))
