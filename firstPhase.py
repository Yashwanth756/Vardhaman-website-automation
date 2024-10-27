import pandas as pd
import numpy as np

class DataCollection:
    def __init__(self, n_chunks, filePath='data.xlsx') -> None:
        file_path = filePath
        data = pd.read_excel(file_path)
        scannedData = self.scanData(data)
        self.divdata = self.divideData(data)

        # Making chunks and resetting index
        for ind, data in enumerate(self.divdata):
            data = self.chunks(data, n_chunks)
            # Reset index for each chunk
            self.divdata[ind] = [chunk.reset_index(drop=True) for chunk in data if not chunk.empty]

    def chunks(self, data, n):
        return np.array_split(data, n)

    def scanData(self, data):
        return data 
    def divideData(self, data):
        return [data]  
