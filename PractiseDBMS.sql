--Ankit Bansal top 10 interview questions
-- https://www.youtube.com/watch?v=Iv9qBz-cyVA&t=23s
-- get the list of emp who is not available in dept table
Select * from emp WHERE emp.department IN (SELECT dept_id FROM dept);

Select * from emp left join dept on department = dept_id where dept_id is NULL;

-- get the 2nd highest sal from each dept
Select * from (SELECT Salary, dense_rank OVER (Partition by department_id order by Salary desc) as rn FROM emp ) a WHERE rn = 2;

-- find all transactions done by shilpa
SELECT * from orders where cust_name = 'shilpa'
--Note: sql server is case insensitive, but in other languages we will get only the exact matching results to counter that we need to use Binary