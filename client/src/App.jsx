import React, { useState } from 'react';
import './App.css';

function App() {
  const [rawFile, setRawFile] = useState(null);
  const [custFile, setCustFile] = useState(null);
  const [filteredUrl, setFilteredUrl] = useState(null);
  const [loading, setLoading] = useState(false);

  const handleRawFileChange = (e) => {
    setRawFile(e.target.files[0]);
    setFilteredUrl(null);
  };

  const handleCustFileChange = (e) => {
    setCustFile(e.target.files[0]);
    setFilteredUrl(null);
  };

  const handleFilter = async () => {
    if (!rawFile || !custFile) return;
    setLoading(true);
    const formData = new FormData();
    formData.append('file', rawFile);
    formData.append('custData', custFile);
    try {
      // use VITE_API_URL env var when deployed; defaults to relative path in dev
      const apiUrl = import.meta.env.VITE_API_URL || '/api/filter';
      const response = await fetch(apiUrl, {
        method: 'POST',
        body: formData,
      });
      if (!response.ok) throw new Error('Failed to filter file');
      const blob = await response.blob();
      setFilteredUrl(URL.createObjectURL(blob));
    } catch (err) {
      console.error('Error filtering file:', err);
      alert('Error filtering file: ' + err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <main className="container">
      <h1 className="header-text">Upload Excel Files</h1>
      <div className='parent'>
      <section className="upload-section">
        <label style={{color: 'white'}}>
          Raw Data File:
          <input type="file" accept=".xlsx" onChange={handleRawFileChange}/>
        </label>
        <label style={{color: 'white'}}>
          Customer Data File :
          <input type="file" accept=".xlsx" onChange={handleCustFileChange}/>
        </label>
        <button onClick={handleFilter} disabled={!rawFile || !custFile || loading}>
          {loading ? 'Filtering...' : 'Filter Data'}
        </button>
      </section>
      {filteredUrl && (
        <section className="download-section">
          <a href={filteredUrl} download="filtered.xlsx">Download File</a>
        </section>
      )}
      </div>
    </main>
  );
}

export default App;
