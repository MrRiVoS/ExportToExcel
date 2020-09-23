using System;

public class DataSet
{
	public DataSet(int size)
	{
		_data = new double[size];
        for (int i = 0; i < size; i++)
        {
			_data[i] = 0.0;
		}
	}

	private double[] _data;

	public void SetData(int pos, double value)
    {
		_data[pos] = value;
    }

	public double GetDataItem(int pos)
    {
		return _data[pos];
    }
}
