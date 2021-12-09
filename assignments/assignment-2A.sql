select
    c.first_name,
    c.last_name,
    sum(f.rental_rate)  as total_earned
from customer c
inner join rental r on r.customer_id = c.customer_id
inner join inventory i on i.inventory_id = r.inventory_id
inner join film f on f.film_id = i.film_id
group by
    c.first_name,
    c.last_name
order by total_earned desc
;
