import { Component, OnInit } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import { SpService } from './services/sp-service';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [RouterOutlet],
  templateUrl: './app.component.html',
  styleUrl: './app.component.scss',
})
export class AppComponent implements OnInit {
  title = 'pms-app';
  isLoading = true;
  constructor(private spService: SpService) {}

  ngOnInit(): void {
    this.spService.getCurrentUser().subscribe((res: any) => {
      this.spService.currentUser = res;
      this.isLoading = false;
    });
  }
}
