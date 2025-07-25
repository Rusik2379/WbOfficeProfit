import React, { useState } from 'react';
import axios from 'axios';

function App() {
  const [file, setFile] = useState(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);
  const [success, setSuccess] = useState(false);

  const handleFileChange = (e) => {
    setFile(e.target.files[0]);
    setSuccess(false);
    setError(null);
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!file) {
      setError('Please select a file');
      return;
    }
  
    setIsLoading(true);
    setError(null);
    setSuccess(false);
  
    const formData = new FormData();
    formData.append('file', file);
  
    try {
      const response = await axios.post('http://localhost:8000/upload/', formData, {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
        responseType: 'blob',
      });
  
      // Check if the response is actually a file
      if (response.data.size === 0) {
        throw new Error('Server returned empty response');
      }
  
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', `processed_${file.name}`);
      document.body.appendChild(link);
      link.click();
      link.parentNode.removeChild(link);
  
      setSuccess(true);
    } catch (err) {
      // Try to read error message if response is JSON
      if (err.response?.data?.text) {
        try {
          const errorText = await err.response.data.text();
          setError(errorText || 'Error processing file');
        } catch {
          setError(err.message || 'Error processing file');
        }
      } else {
        setError(err.message || 'Error processing file');
      }
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className="App" style={{ padding: '20px', maxWidth: '600px', margin: '0 auto' }}>
      <h1>Обработка финансовых отчетов</h1>
      <p>Загрузите Excel файл для применения макросов</p>
      <p><small>Файл будет обработан: сначала фильтрация строк, затем расчет прибыли</small></p>
      
      <form onSubmit={handleSubmit}>
        <div style={{ margin: '20px 0' }}>
          <input 
            type="file" 
            accept=".xls,.xlsx" 
            onChange={handleFileChange} 
            style={{ display: 'block', margin: '10px 0' }}
          />
        </div>
        <button 
          type="submit" 
          disabled={isLoading}
          style={{ 
            padding: '10px 20px', 
            background: '#4CAF50', 
            color: 'white', 
            border: 'none', 
            borderRadius: '4px',
            cursor: 'pointer',
            fontSize: '16px'
          }}
        >
          {isLoading ? 'Обработка...' : 'Обработать файл'}
        </button>
      </form>

      {error && (
        <div style={{ 
          marginTop: '20px', 
          padding: '15px', 
          background: '#ffebee', 
          color: '#d32f2f',
          borderRadius: '4px'
        }}>
          {error}
        </div>
      )}

      {success && (
        <div style={{ 
          marginTop: '20px', 
          padding: '15px', 
          background: '#e8f5e9', 
          color: '#2e7d32',
          borderRadius: '4px'
        }}>
          Файл успешно обработан и загружается!
        </div>
      )}
    </div>
  );
}

export default App;