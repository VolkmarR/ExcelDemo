type HomeProps = {
    selectedFile: File | null
    onFileSelected: (file: File | null) => void
}

function Home({ selectedFile, onFileSelected }: HomeProps) {
    return (
        <div className="home">
            <h2>Select an Excel file</h2>

            <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => {
                    const file = e.target.files?.[0] ?? null
                    onFileSelected(file)
                }}
            />

            {selectedFile && (
                <p>
                    Selected: <strong>{selectedFile.name}</strong>
                </p>
            )}
        </div>
    )
}

export default Home