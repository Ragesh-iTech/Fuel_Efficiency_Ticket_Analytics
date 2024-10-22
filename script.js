
let vehicles = [];

document.getElementById('upload').addEventListener('change', handleFileUpload);

function handleFileUpload(event) {
    const file = event.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

            vehicles = [];
            rows.slice(1).forEach(row => {
                const vehicleNo = row[0];
                const distance = parseFloat(row[1]);
                const dieselConsumed = parseFloat(row[2]);
                const ticketSales = parseFloat(row[3]);

                const vehicle = {
                    vehicleNo,
                    distance,
                    dieselConsumed,
                    ticketSales,
                    efficiency: dieselConsumed > 0 ? (distance / dieselConsumed).toFixed(2) : 0
                };

                vehicles.push(vehicle);
            });

            displayVehicles();
        };
        reader.readAsArrayBuffer(file);
    }
}

function displayVehicles() {
    const vehicleList = document.getElementById("vehicleList");
    vehicleList.innerHTML = "";

    vehicles.forEach(vehicle => {
        const li = document.createElement("li");
        li.textContent = `Vehicle No: ${vehicle.vehicleNo}, Distance: ${vehicle.distance} km, Diesel Consumed: ${vehicle.dieselConsumed} L, Efficiency: ${vehicle.efficiency} km/l, Ticket Sales: $${vehicle.ticketSales}`;
        vehicleList.appendChild(li);
    });
}

function finalizeData() {
    if (vehicles.length === 0) {
        alert("No vehicles to analyze.");
        return;
    }

    const totalFuel = calculateTotalFuel();
    const totalTicketSales = calculateTotalTicketSales();
    const avgEfficiency = calculateAverageEfficiency();
    const bestVehicle = getBestVehicle();
    const worstVehicle = getWorstVehicle();

    displayAnalysis(totalFuel, totalTicketSales, avgEfficiency, bestVehicle, worstVehicle);
    displayCharts();
}

function calculateTotalFuel() {
    return vehicles.reduce((total, vehicle) => total + vehicle.dieselConsumed, 0).toFixed(2);
}

function calculateTotalTicketSales() {
    return vehicles.reduce((total, vehicle) => total + vehicle.ticketSales, 0).toFixed(2);
}

function calculateAverageEfficiency() {
    const totalEfficiency = vehicles.reduce((total, vehicle) => total + parseFloat(vehicle.efficiency), 0);
    return (totalEfficiency / vehicles.length).toFixed(2);
}

function getBestVehicle() {
    return vehicles.reduce((best, vehicle) => (parseFloat(vehicle.efficiency) > parseFloat(best.efficiency) ? vehicle : best), vehicles[0]);
}

function getWorstVehicle() {
    return vehicles.reduce((worst, vehicle) => (parseFloat(vehicle.efficiency) < parseFloat(worst.efficiency) ? vehicle : worst), vehicles[0]);
}

function displayAnalysis(totalFuel, totalTicketSales, avgEfficiency, bestVehicle, worstVehicle) {
    document.getElementById("totalFuel").textContent = `Total Diesel Consumed: ${totalFuel} liters`;
    document.getElementById("totalTicketSales").textContent = `Total Ticket Sales: $${totalTicketSales}`;
    document.getElementById("averageEfficiency").textContent = `Average Efficiency: ${avgEfficiency} km/l`;
    document.getElementById("bestVehicle").textContent = `Best Vehicle: ${bestVehicle.vehicleNo} with Efficiency of ${bestVehicle.efficiency} km/l`;
    document.getElementById("worstVehicle").textContent = `Worst Vehicle: ${worstVehicle.vehicleNo} with Efficiency of ${worstVehicle.efficiency} km/l`;
}

function displayCharts() {
    // Efficiency Line Graph
    const efficiencyCtx = document.getElementById('efficiencyChart').getContext('2d');
    const vehicleLabels = vehicles.map(vehicle => vehicle.vehicleNo);
    const efficiencies = vehicles.map(vehicle => parseFloat(vehicle.efficiency));

    new Chart(efficiencyCtx, {
        type: 'line',
        data: {
            labels: vehicleLabels,
            datasets: [{
                label: 'Fuel Efficiency (km/l)',
                data: efficiencies,
                borderColor: 'rgba(75, 192, 192, 1)',
                fill: false,
                tension: 0.1
            }]
        },
        options: {
            scales: {
                y: {
                    beginAtZero: true
                }
            }
        }
    });

    // Ticket Sales Histogram
    const ticketSalesCtx = document.getElementById('ticketSalesChart').getContext('2d');
    const ticketSales = vehicles.map(vehicle => vehicle.ticketSales);

    new Chart(ticketSalesCtx, {
        type: 'bar',
        data: {
            labels: vehicleLabels,
            datasets: [{
                label: 'Ticket Sales ($)',
                data: ticketSales,
                backgroundColor: 'rgba(255, 99, 132, 0.2)',
                borderColor: 'rgba(255, 99, 132, 1)',
                borderWidth: 1
            }]
        },
        options: {
            scales: {
                y: {
                    beginAtZero: true
                }
            }
        }
    });
}


function downloadCSV() {
    const rows = [
        ['Vehicle No', 'Distance (km)', 'Diesel Consumed (L)', 'Efficiency (km/l)', 'Ticket Sales ($)'],
        ...vehicles.map(vehicle => [vehicle.vehicleNo, vehicle.distance, vehicle.dieselConsumed, vehicle.efficiency, vehicle.ticketSales]),
        [],
        ['Total Diesel Consumed:', calculateTotalFuel()],
        ['Total Ticket Sales:', calculateTotalTicketSales()],
        ['Average Efficiency:', calculateAverageEfficiency()],
        ['Best Vehicle:', getBestVehicle().vehicleNo, 'Efficiency:', getBestVehicle().efficiency],
        ['Worst Vehicle:', getWorstVehicle().vehicleNo, 'Efficiency:', getWorstVehicle().efficiency]
    ];

    let csvContent = "data:text/csv;charset=utf-8," + rows.map(row => row.join(",")).join("\n");

    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", "vehicle_analysis.csv");
    document.body.appendChild(link);
    link.click();
}


//copy text on clipboard
// JavaScript
function copyToClipboard() {
    const textToCopy = document.querySelector(".para");
    navigator.clipboard.writeText(textToCopy.textContent)
        .then(() => alert('Text copied!'))
        .catch(err => console.error('Failed to copy text: ', err));
}

// Function to download efficiency chart
function downloadEfficiencyImage() {
    const canvas = document.getElementById('efficiencyChart');
    
    // Convert the canvas to a data URL
    const dataURL = canvas.toDataURL('image/png');
    
    // Create a temporary link element
    const downloadLink = document.createElement('a');
    downloadLink.href = dataURL;
    downloadLink.download = 'efficiency_chart.png'; // Set the download filename
    
    // Append link to body, click it, and remove it
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
}

// Function to download ticket sales chart
function downloadTicketSalesImage() {
    const canvas = document.getElementById('ticketSalesChart');
    
    // Convert the canvas to a data URL
    const dataURL = canvas.toDataURL('image/png');
    
    // Create a temporary link element
    const downloadLink = document.createElement('a');
    downloadLink.href = dataURL;
    downloadLink.download = 'ticket_sales_chart.png'; // Set the download filename
    
    // Append link to body, click it, and remove it
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
}


