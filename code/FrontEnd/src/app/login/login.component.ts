import { Component } from '@angular/core';
import { Router } from '@angular/router';

@Component({
  selector: 'app-login',
  templateUrl: './login.component.html',
  styleUrls: ['./login.component.css']
})
export class LoginComponent {
  email: string = '';
  password: string = '';
  emailError: boolean = false;
  passwordError: boolean = false;

  constructor(private router: Router) { }

  login() {
    this.emailError = !this.email;
    this.passwordError = !this.password;
    
    if (this.email && this.password) {
      // Check credentials - using admin/admin as per the memory
      if (this.email === 'admin' && this.password === 'admin') {
        // Store username in localStorage to display on welcome page
        localStorage.setItem('username', this.email);
        // Navigate to welcome page
        this.router.navigate(['/welcome']);
      } else {
        alert('Invalid credentials. Please try again.');
      }
    }
  }


}
