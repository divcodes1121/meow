import React from 'react';
import {
  LineChart, Line, BarChart, Bar, PieChart, Pie, Cell,
  AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip,
  Legend, ResponsiveContainer
} from 'recharts';

// Color palette for consistent colors across chart types
const COLORS = [
  '#8884d8', '#82ca9d', '#ffc658', '#ff8042', '#0088FE', 
  '#00C49F', '#FFBB28', '#FF8042', '#a4de6c', '#d0ed57'
];

// Line Chart Component
export const LineChartComponent = ({ data, xKey, yKeys, title }) => {
  return (
    <div className="chart-container">
      <h3 className="chart-title">{title}</h3>
      <ResponsiveContainer width="100%" height="100%">
        <LineChart
          data={data}
          margin={{ top: 5, right: 30, left: 20, bottom: 5 }}
        >
          <CartesianGrid strokeDasharray="3 3" />
          <XAxis dataKey={xKey} />
          <YAxis />
          <Tooltip />
          <Legend />
          {yKeys.map((key, index) => (
            <Line
              key={key}
              type="monotone"
              dataKey={key}
              stroke={COLORS[index % COLORS.length]}
              activeDot={{ r: 8 }}
            />
          ))}
        </LineChart>
      </ResponsiveContainer>
    </div>
  );
};

// Bar Chart Component
export const BarChartComponent = ({ data, xKey, yKeys, title, stacked }) => {
  return (
    <div className="chart-container">
      <h3 className="chart-title">{title}</h3>
      <ResponsiveContainer width="100%" height="100%">
        <BarChart
          data={data}
          margin={{ top: 5, right: 30, left: 20, bottom: 5 }}
        >
          <CartesianGrid strokeDasharray="3 3" />
          <XAxis dataKey={xKey} />
          <YAxis />
          <Tooltip />
          <Legend />
          {yKeys.map((key, index) => (
            <Bar
              key={key}
              dataKey={key}
              fill={COLORS[index % COLORS.length]}
              stackId={stacked ? "stack" : null}
            />
          ))}
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
};

// Pie Chart Component
export const PieChartComponent = ({ data, nameKey, valueKey, title }) => {
  // For pie charts, we need to transform the data to a format that works with recharts
  const transformData = () => {
    // If data is already in correct format
    if (data && data.length > 0 && data[0].hasOwnProperty(nameKey) && data[0].hasOwnProperty(valueKey)) {
      return data;
    }
    
    // Otherwise, create pie chart data
    // Assuming first data item has all necessary keys
    if (!data || data.length === 0) return [];
    
    const firstItem = data[0];
    const pieData = [];
    
    // Using the first row keys as names, and values as values
    Object.keys(firstItem).forEach(key => {
      if (key !== nameKey) {
        pieData.push({
          name: key,
          value: Number(firstItem[key]) || 0
        });
      }
    });
    
    return pieData;
  };
  
  const pieData = transformData();
  
  return (
    <div className="chart-container">
      <h3 className="chart-title">{title}</h3>
      <ResponsiveContainer width="100%" height="100%">
        <PieChart>
          <Pie
            data={pieData}
            cx="50%"
            cy="50%"
            labelLine={true}
            label={({ name, percent }) => `${name}: ${(percent * 100).toFixed(0)}%`}
            outerRadius={80}
            fill="#8884d8"
            dataKey="value"
          >
            {pieData.map((entry, index) => (
              <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
            ))}
          </Pie>
          <Tooltip formatter={(value) => Math.round(value * 100) / 100} />
          <Legend />
        </PieChart>
      </ResponsiveContainer>
    </div>
  );
};

// Area Chart Component
export const AreaChartComponent = ({ data, xKey, yKeys, title, stacked }) => {
  return (
    <div className="chart-container">
      <h3 className="chart-title">{title}</h3>
      <ResponsiveContainer width="100%" height="100%">
        <AreaChart
          data={data}
          margin={{ top: 5, right: 30, left: 20, bottom: 5 }}
        >
          <CartesianGrid strokeDasharray="3 3" />
          <XAxis dataKey={xKey} />
          <YAxis />
          <Tooltip />
          <Legend />
          {yKeys.map((key, index) => (
            <Area
              key={key}
              type="monotone"
              dataKey={key}
              stackId={stacked ? "1" : index}
              stroke={COLORS[index % COLORS.length]}
              fill={COLORS[index % COLORS.length]}
              fillOpacity={0.6}
            />
          ))}
        </AreaChart>
      </ResponsiveContainer>
    </div>
  );
};

export default {
  LineChartComponent,
  BarChartComponent,
  PieChartComponent,
  AreaChartComponent
};