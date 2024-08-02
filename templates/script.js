document.addEventListener('DOMContentLoaded', function () {
    fetch('data.json')
        .then(response => {
            if (!response.ok) {
                throw new Error('Network response was not ok');
            }
            return response.json();
        })
        .then(data => {
            document.getElementById('text_content').innerHTML = data.text;
            document.getElementById('greetings').innerHTML = data.greetings.replace(/\n/g, '<br>');
            document.getElementById('table_of_contents').innerHTML = data.table_of_contents.replace(/\n/g, '<br>');
        })
        .catch(error => console.error('Error loading JSON:', error));
});



document.addEventListener('DOMContentLoaded', function() {
    var ctx = document.getElementById('donutChart').getContext('2d');
    
    var donutChart = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels: ['Total Expenditure', 'Profit'],
            datasets: [{
                label: 'OVERVIEEW',
                data: [139493.05, 51895.79],
                borderWidth: 1
            }]
        },
       
    });
});
document.addEventListener('DOMContentLoaded', function () {
    var ctx = document.getElementById('myHorizontalBarChart').getContext('2d');

    var myHorizontalBarChart = new Chart(ctx, {
        type: 'bar', // Chart type
        data: {
            labels: ['Actual Revenue', 'Budgeted Revenue', 'Previous Period Revenue'],
            datasets: [{
                label: 'Revenue Growth',
                data: [191388.84, 168224.66, 223980.90], // Corrected data array
                backgroundColor: ['rgba(75, 192, 192, 0.2)', 'rgba(153, 102, 255, 0.2)', 'rgba(255, 159, 64, 0.2)'], // Added color for the third bar
                borderColor: ['rgba(75, 192, 192, 1)', 'rgba(153, 102, 255, 1)', 'rgba(255, 159, 64, 1)'], // Added color for the third bar
                borderWidth: 1
            }]
        },
        options: {
            indexAxis: 'y', // This makes the bars horizontal
            scales: {
                x: {
                    beginAtZero: true // Ensures that the x-axis starts at 0
                },
                y: {
                    beginAtZero: true // Ensures that the y-axis starts at 0
                }
            }
        }
    });
});