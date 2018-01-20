package com.imdbtopmovies;

public class Movie {

    public String Rank;
    public String OldRank;
    public String Name;
    public String Year;
    public String Status;

    public Movie() {
    }

    public Movie(String rank, String name, String year) {
        this.Rank = rank;
        this.Name = name;
        this.Year = year;

    }

    public Movie(String rank, String oldRank, String name, String year, String status) {
        this.Rank = rank;
        this.OldRank = oldRank;
        this.Name = name;
        this.Year = year;
        this.Status = status;
    }
}
